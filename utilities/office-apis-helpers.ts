import { AxiosResponse } from 'axios';
import { PublicClientApplication, AccountInfo } from '@azure/msal-browser';
import { registerTokenRenewal } from './microsoft-graph-helpers';
import { AppState } from '../src/components/App';

/*
     Interacting with the Office document
*/
export const writeFileNamesToWorksheet = async (
  result: AxiosResponse,
  displayError: (x: string) => void
) => {
  try {
    await Excel.run((context: Excel.RequestContext) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const data = [
        [result.data.value[0].name],
        [result.data.value[1].name],
        [result.data.value[2].name],
      ];
      const range = sheet.getRange('B5:B7');
      range.values = data;
      range.format.autofitColumns();
      return context.sync();
    });
  } catch (error) {
    displayError(error.toString());
  }
};

/*
    Managing the dialogs and authentication - VERSIÓN SIMPLIFICADA
*/

interface DialogEventReceivedArgs {
  error: number;
  type: string;
}

// Variables globales
let loginDialog: Office.Dialog;
let msalInstance: PublicClientApplication | null = null;
let currentUser: AccountInfo | null = null;
let currentToken: string = '';

const dialogLoginUrl: string =
  location.protocol + '//' + location.hostname + 
  (location.port ? ':' + location.port : '') + '/login/login.html';

// Crear configuración MSAL única
const createMsalConfig = () => {
  const sessionId = Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  const userContext = Office.context.document?.url || 'unknown';
  const userHash = btoa(userContext).substr(0, 8);
  
  return {
    auth: {
      clientId: 'e3f4535d-6cc9-4ba7-ae00-3a3fffecfdd7',
      authority: 'https://login.microsoftonline.com/26870be5-0feb-4774-a29a-9bee84427a21',
      redirectUri: dialogLoginUrl,
    },
    cache: {
      cacheLocation: 'sessionStorage' as const,
      storeAuthStateInCookie: false,
      cacheName: `msal_cache_${userHash}_${sessionId}`
    },
    system: {
      allowNativeBroker: false,
      loggerOptions: {
        loggerCallback: (level, message, containsPii) => {
          if (!containsPii) {
            console.log(`[MSAL-${level}] ${message}`);
          }
        },
        piiLoggingEnabled: false,
        logLevel: 3
      }
    }
  };
};

// Inicializar MSAL
const initializeMsal = async (): Promise<PublicClientApplication> => {
  if (!msalInstance) {
    const config = createMsalConfig();
    msalInstance = new PublicClientApplication(config);
    await msalInstance.initialize();
  }
  return msalInstance;
};

// Función para validar token
const validateToken = async (token: string): Promise<boolean> => {
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` }
    });
    const isValid = response.ok;
    return isValid;
  } catch (error) {
    console.warn('Error validando token:', error);
    return false;
  }
};

// Función para renovar token - EXPORTADA para usar en graph-helpers
const renewTokenInternal = async (): Promise<string | null> => {
  try {
    if (!msalInstance || !currentUser) {
      console.warn('⚠️ No hay instancia MSAL o usuario para renovar token');
      return null;
    }
    
    const silentRequest = {
      scopes: ['User.Read', 'Files.ReadWrite.All', 'Sites.ReadWrite.All'],
      account: currentUser,
      forceRefresh: false
    };

    const response = await msalInstance.acquireTokenSilent(silentRequest);
    const newToken = response.accessToken;
    
    // Actualizar token actual
    currentToken = newToken;
    
    return newToken;
  } catch (error) {
    console.warn('⚠️ No se pudo renovar token:', error);
    return null;
  }
};

const setupTokenRenewal = () => {
  registerTokenRenewal(renewTokenInternal);
};

// FUNCIÓN PRINCIPAL DE LOGIN
export const signInO365 = async (
  setState: (x: AppState) => void,
  setToken: (x: string) => void,
  setUserName: (x: string) => void,
  displayError: (x: string) => void
) => {
  setState({ authStatus: 'loginInProcess' });

  try {
    // Inicializar MSAL
    const pca = await initializeMsal();
    setupTokenRenewal();
    try {
      const ssoToken = await Office.auth.getAccessToken({
        allowSignInPrompt: false,
        allowConsentPrompt: false,
        forMSGraphAccess: true,
      });

      const isValidSSO = await validateToken(ssoToken);
      if (isValidSSO) {
        const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: { Authorization: `Bearer ${ssoToken}` }
        });
        
        if (graphResponse.ok) {
          const userData = await graphResponse.json();
          currentToken = ssoToken;
          setToken(ssoToken);
          setUserName(userData.userPrincipalName || '');
          setState({
            authStatus: 'loggedIn',
            headerMessage: 'Bienvenido ' + (userData.displayName || userData.userPrincipalName),
          });
          return;
        }
      }
    } catch (ssoError) {
      console.warn('SSO falló:', ssoError);
    }

    const accounts = pca.getAllAccounts();
    
    if (accounts.length > 0) {
      const account = accounts[0];
      currentUser = account;
      const renewedToken = await renewTokenInternal();
      
      if (renewedToken) {
        const isValid = await validateToken(renewedToken);
        if (isValid) {
          currentToken = renewedToken;
          setToken(renewedToken);
          setUserName(account.username || '');
          setState({
            authStatus: 'loggedIn',
            headerMessage: 'Bienvenido ' + (account.name || account.username),
          });
          return;
        }
      }
    }
    await startInteractiveLogin(setState, setToken, setUserName, displayError);
  } catch (error) {
    console.error('Error en signInO365:', error);
    displayError(`Error al iniciar sesión: ${error.message}`);
    setState({ authStatus: 'notLoggedIn' });
  }
};

// Login interactivo
const startInteractiveLogin = async (
  setState: (x: AppState) => void,
  setToken: (x: string) => void,
  setUserName: (x: string) => void,
  displayError: (x: string) => void
) => {
  Office.context.ui.displayDialogAsync(
    dialogLoginUrl,
    { height: 40, width: 30 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('Error abriendo diálogo:', result.error);
        displayError(`${result.error.code} ${result.error.message}`);
        setState({ authStatus: 'notLoggedIn' });
      } else {
        loginDialog = result.value;
        
        loginDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: any) => processLoginMessage(arg, setState, setToken, setUserName, displayError)
        );
        
        loginDialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (arg: DialogEventReceivedArgs) => processLoginDialogEvent(arg, setState, displayError)
        );
      }
    }
  );
};

// Procesar mensaje del diálogo
const processLoginMessage = async (
  arg: any,
  setState: (x: AppState) => void,
  setToken: (x: string) => void,
  setUserName: (x: string) => void,
  displayError: (x: string) => void
) => {
  if (arg.origin !== window.location.origin) {
    console.error('Origen incorrecto');
    displayError('Origen incorrecto en el mensaje del diálogo.');
    setState({ authStatus: 'notLoggedIn' });
    return;
  }

  let messageFromDialog;
  try {
    messageFromDialog = JSON.parse(arg.message);
  } catch (error) {
    console.error('Error parseando mensaje:', error);
    displayError('Error al procesar el mensaje del diálogo.');
    setState({ authStatus: 'notLoggedIn' });
    return;
  }

  if (messageFromDialog.status === 'success') {
    try {
      const isValid = await validateToken(messageFromDialog.token);
      
      if (isValid) {
        // Cerrar diálogo
        if (loginDialog) loginDialog.close();

        // Actualizar usuario actual
        if (msalInstance) {
          const accounts = msalInstance.getAllAccounts();
          if (accounts.length > 0) {
            currentUser = accounts[0];
          }
        }

        // Establecer token actual
        currentToken = messageFromDialog.token;
        setToken(messageFromDialog.token);
        setUserName(messageFromDialog.userName || '');
        setState({
          authStatus: 'loggedIn',
          headerMessage: 'Bienvenido ' + (messageFromDialog.userName || 'Usuario'),
        });
      } else {
        throw new Error('Token recibido no es válido');
      }
    } catch (error) {
      console.error('❌ Error procesando token:', error);
      if (loginDialog) loginDialog.close();
      displayError('Error validando credenciales. Intente nuevamente.');
      setState({ authStatus: 'notLoggedIn' });
    }
  } else {
    console.error('❌ Error en login:', messageFromDialog.result);
    if (loginDialog) loginDialog.close();
    displayError(messageFromDialog.result?.errorMessage || 'Error desconocido');
    setState({ authStatus: 'notLoggedIn' });
  }
};

export const getValidToken = async (): Promise<string | null> => {
  if (currentToken) {
    const isValid = await validateToken(currentToken);
    if (isValid) {
      return currentToken;
    }
  }
  const renewedToken = await renewTokenInternal();
  
  if (renewedToken) {
    return renewedToken;
  }
  return null;
};

// Función para obtener token actual (para uso directo)
export const getCurrentToken = (): string => {
  return currentToken;
};

// Limpiar recursos
export const cleanupAuth = () => {
  currentUser = null;
  currentToken = '';
};

// LOGOUT
export const logoutFromO365 = async (
  setState: (x: AppState) => void,
  setUserName: (x: string) => void,
  _userName: string,
  _displayError: (x: string) => void
) => {
  try {
    cleanupAuth();
    
    if (msalInstance && currentUser) {
      await msalInstance.logout({ account: currentUser });
    }
    
    setState({
      authStatus: 'notLoggedIn',
      headerMessage: 'Welcome',
    });
    setUserName('');
    
  } catch (error) {
    setState({
      authStatus: 'notLoggedIn',
      headerMessage: 'Welcome',
    });
    setUserName('');
  }
};

// Resto de funciones sin cambios...
const processLoginDialogEvent = (
  arg: DialogEventReceivedArgs,
  setState: (x: AppState) => void,
  displayError: (x: string) => void
) => {
  processDialogEvent(arg, setState, displayError);
};

const processDialogEvent = (
  arg: DialogEventReceivedArgs,
  setState: (x: AppState) => void,
  displayError: (x: string) => void
) => {
  const errorCode = arg.error;
  switch (errorCode) {
    case 12002:
      displayError('El diálogo no pudo cargar la página especificada.');
      break;
    case 12003:
      displayError('El diálogo requiere HTTPS.');
      break;
    case 12006:
      setState({
        authStatus: 'notLoggedIn',
        headerMessage: 'Welcome',
      });
      break;
    default:
      displayError('Error desconocido en el diálogo.');
      break;
  }
};