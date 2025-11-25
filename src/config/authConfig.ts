//import { AppState } from '../src/components/App';
import { AxiosResponse } from 'axios';
import { PublicClientApplication } from "@azure/msal-browser";
import { AppState } from '../components/App';

/*
    MSAL Configuration
*/
export const msalConfig = {
  auth: {
    clientId: 'e3f4535d-6cc9-4ba7-ae00-3a3fffecfdd7',
    authority: 'https://login.microsoftonline.com/26870be5-0feb-4774-a29a-9bee84427a21',
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: ["User.Read", "Files.Read.All"]
};

/*
    Check for existing authentication - Multi-user support
*/
export const checkExistingAuth = async (
  setState: (x: AppState) => void,
  setToken: (x: string) => void,
  setUserName: (x: string) => void,
  _displayError: (x: string) => void
): Promise<boolean> => {
  try {
    const pca = new PublicClientApplication(msalConfig);
    await pca.initialize();
    
    // Get all accounts - this allows multiple users on the same machine
    const accounts = pca.getAllAccounts();
    
    if (accounts.length > 0) {
      // For multi-user support, we'll use the first account found
      // In a production app, you might want to let users choose which account to use
      const account = accounts[0];
      
      try {
        // Try to acquire token silently for this specific account
        const silentRequest = {
          ...loginRequest,
          account: account,
          forceRefresh: false
        };
        
        const response = await pca.acquireTokenSilent(silentRequest);
        
        if (response && response.accessToken) {
          // We have a valid token, user is already authenticated
          setToken(response.accessToken);
          setUserName(response.account?.username || account.username);
          setState({
            authStatus: 'loggedIn',
            headerMessage: `Bienvenido ${response.account?.name || response.account?.username}`
          });
          return true;
        }
      } catch (silentError) {
        // Silent token acquisition failed, might be expired or revoked
        // Silent token acquisition failed
        // displayError((silentError as Error).message || 'Silent authentication failed. Please log in again.');
        // If it's an interaction required error, we'll need interactive login
        if (silentError.errorCode === 'interaction_required' || 
            silentError.errorCode === 'consent_required' ||
            silentError.errorCode === 'login_required') {
          // Interactive login required
        }
      }
    }
    
    // No existing session found or silent auth failed
    return false;
    
  } catch (error) {
    console.error('Error checking existing auth:', error);
    // Don't show error to user for auth check failures, just proceed to login
    return false;
  }
};

/*
    Initialize authentication on app start
*/
export const initializeAuth = async (
  setState: (x: AppState) => void,
  setToken: (x: string) => void,
  setUserName: (x: string) => void,
  displayError: (x: string) => void
) => {
  // First check if user is already authenticated
  const isAuthenticated = await checkExistingAuth(setState, setToken, setUserName, displayError);
  
  if (!isAuthenticated) {
    // User is not authenticated, set initial state
    setState({
      authStatus: 'notLoggedIn',
      headerMessage: 'Welcome'
    });
  }
};

/*
     Interacting with the Office document
*/
export const writeFileNamesToWorksheet = async (result: AxiosResponse,
    displayError: (x: string) => void) => {

    try {
        await Excel.run((context: Excel.RequestContext) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            const data = [
                [result.data.value[0].name],
                [result.data.value[1].name],
                [result.data.value[2].name]];

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
    Managing the dialogs - Updated to check existing auth first
*/

let loginDialog: Office.Dialog;
const dialogLoginUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/login/login.html';

export const signInO365 = async (setState: (x: AppState) => void,
    setToken: (x: string) => void,
    setUserName: (x: string) => void,
    displayError: (x: string) => void) => {

    // First check if user is already authenticated
    const isAuthenticated = await checkExistingAuth(setState, setToken, setUserName, displayError);
    
    if (isAuthenticated) {
        // User is already logged in, no need to show dialog
        return;
    }

    // User needs to authenticate
    setState({ authStatus: 'loginInProcess' });

    Office.context.ui.displayDialogAsync(
        dialogLoginUrl,
        { height: 40, width: 30 },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                displayError(`${result.error.code} ${result.error.message}`);
            }
            else {
                loginDialog = result.value;
                loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginMessage);
                loginDialog.addEventHandler(Office.EventType.DialogEventReceived, processLoginDialogEvent);
            }
        }
    );

    const processLoginMessage = (arg: { message: string, origin: string }) => {
        // Confirm origin is correct.
        if (arg.origin !== window.location.origin) {
            throw new Error("Incorrect origin passed to processLoginMessage.");
        }

        let messageFromDialog = JSON.parse(arg.message);
        if (messageFromDialog.status === 'success') {

            // We now have a valid access token.
            loginDialog.close();
            setToken(messageFromDialog.token);
            setUserName(messageFromDialog.userName);
            setState({
                authStatus: 'loggedIn',
                headerMessage: 'Bienvenido ' + messageFromDialog.userName
            });
        }
        else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            displayError(messageFromDialog.result);
        }
    };

    const processLoginDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

let logoutDialog: Office.Dialog;
const dialogLogoutUrl: string = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + '/logout/logout.html';

// From https://stackoverflow.com/questions/37764665/how-to-implement-sleep-function-in-typescript
function delay(milliSeconds: number) {
    return new Promise(resolve => setTimeout(resolve, milliSeconds));
}

export const logoutFromO365 = async (setState: (x: AppState) => void,
    setUserName: (x: string) => void,
    setToken: (x: string) => void,
    userName: string,
    displayError: (x: string) => void) => {

    try {
        // Clear MSAL cache for current user
        const pca = new PublicClientApplication(msalConfig);
        await pca.initialize();
        
        const accounts = pca.getAllAccounts();
        
        // Find the specific account to logout (important for multi-user scenarios)
        const currentAccount = accounts.find(account => 
            account.username === userName || 
            account.name === userName
        );
        
        if (currentAccount) {
            // Log out the current user's account
            await pca.logout({ account: currentAccount, postLogoutRedirectUri: window.location.origin });
        } else if (accounts.length > 0) {
            // Fallback: log out the first account if we can't find the specific one
            await pca.logout({ account: accounts[0], postLogoutRedirectUri: window.location.origin });
        }
        
        // Note: We don't clear ALL localStorage since other users might be cached
        // MSAL handles this internally per account
        
    } catch (error) {
        console.error('Error during logout:', error);
    }

    Office.context.ui.displayDialogAsync(dialogLogoutUrl,
        { height: 40, width: 30 },
        async (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                displayError(`${result.error.code} ${result.error.message}`);
            }
            else {
                logoutDialog = result.value;
                logoutDialog.addEventHandler(Office.EventType.DialogMessageReceived, processLogoutMessage);
                logoutDialog.addEventHandler(Office.EventType.DialogEventReceived, processLogoutDialogEvent);
                await delay(5000); // Wait for dialog to initialize and register handler for messaging.
                logoutDialog.messageChild(JSON.stringify({ "userName": userName }));
            }
        }
    );

    const processLogoutMessage = () => {
        logoutDialog.close();
        setState({
            authStatus: 'notLoggedIn',
            headerMessage: 'Welcome'
        });
        setUserName('');
        setToken('');
    };

    const processLogoutDialogEvent = (arg) => {
        processDialogEvent(arg, setState, displayError);
    };
};

const processDialogEvent = (arg: { error: number, type: string },
    setState: (x: AppState) => void,
    displayError: (x: string) => void) => {

    switch (arg.error) {
        case 12002:
            displayError('The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.');
            break;
        case 12003:
            displayError('The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.');
            break;
        case 12006:
            // 12006 means that the user closed the dialog instead of waiting for it to close.
            // It is not known if the user completed the login or logout, so assume the user is
            // logged out and revert to the app's starting state. It does no harm for a user to
            // press the login button again even if the user is logged in.
            setState({
                authStatus: 'notLoggedIn',
                headerMessage: 'Welcome'
            });
            break;
        default:
            displayError('Unknown error in dialog box.');
            break;
    }
};