/*
 * Updated login.html script to better handle existing sessions
 */
import { PublicClientApplication } from "@azure/msal-browser";

Office.onReady(async () => {
  const pca = new PublicClientApplication({
    auth: {
      clientId: 'e3f4535d-6cc9-4ba7-ae00-3a3fffecfdd7',
      authority: 'https://login.microsoftonline.com/26870be5-0feb-4774-a29a-9bee84427a21',
      redirectUri: `${window.location.origin}/login/login.html`
    },
    cache: {
      cacheLocation: 'localStorage'
    }
  });
  
  await pca.initialize();
  
  try {
    // First, handle any redirect response
    const response = await pca.handleRedirectPromise();
    
    if (response) {
      // Login was successful via redirect
      Office.context.ui.messageParent(JSON.stringify({ 
        status: 'success', 
        token: response.accessToken, 
        userName: response.account.username
      }));
      return;
    }
    
    // Check if user is already logged in
    const accounts = pca.getAllAccounts();
    
    if (accounts.length > 0) {
      // Try to get token silently
      try {
        const silentRequest = {
          scopes: ['user.read', 'files.read.all'],
          account: accounts[0],
          forceRefresh: false
        };
        
        const silentResponse = await pca.acquireTokenSilent(silentRequest);
        
        if (silentResponse && silentResponse.accessToken) {
          // We have a valid token
          Office.context.ui.messageParent(JSON.stringify({ 
            status: 'success', 
            token: silentResponse.accessToken, 
            userName: silentResponse.account.username
          }));
          return;
        }
      } catch (silentError) {
        console.log('Silent token acquisition failed:', silentError);
        // Fall through to interactive login
      }
    }
    
    // No existing session or silent auth failed, need interactive login
    await pca.loginRedirect({
      scopes: ['user.read', 'files.read.all']
    });
    
  } catch (error) {
    const errorData = {
      errorMessage: error.errorCode,
      message: error.errorMessage,
      errorCode: error.stack
    };
    Office.context.ui.messageParent(JSON.stringify({ 
      status: 'failure', 
      result: errorData 
    }));
  }
});