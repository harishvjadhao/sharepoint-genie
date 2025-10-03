import { PublicClientApplication, InteractionRequiredAuthError, AccountInfo } from '@azure/msal-browser';

export interface IAuthSettings {
  appClientId: string;
  tenantId: string;
  currentUserLogin?: string;
  redirectUri?: string;  
}

export async function acquireToken(settings: IAuthSettings): Promise<string | undefined> {
  if (!settings?.appClientId || !settings?.tenantId) {
    return undefined;
  }

  console.log('Using redirect URI:', settings.redirectUri || window.location.origin);

  const msalInstance = new PublicClientApplication({
    auth: {
      clientId: settings.appClientId,
      authority: `https://login.microsoftonline.com/${settings.tenantId}`,
      redirectUri: settings.redirectUri || window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
    }
  });

  await msalInstance.initialize();

  const scopes = ['api://a76fed20-34ae-4be6-88f3-b2b9f3ac78f5/access_as_user']; // add Graph scopes you need

  // get cached account if available
  const accounts = msalInstance.getAllAccounts();
  let account: AccountInfo | null = accounts && accounts.length > 0 ? accounts[0] : null;

  // if no cached account → do interactive login via popup
  if (!account) {
    try {
      const loginResponse = await msalInstance.loginPopup({
        scopes,
        loginHint: settings.currentUserLogin,
      });
      account = loginResponse.account || null;
      console.log("User logged in via popup:", account?.username);
    } catch (err) {
      console.error("loginPopup failed:", err);
      return undefined;
    }
  }

  // try to get token silently
  try {
    const response = await msalInstance.acquireTokenSilent({
      scopes,
      account,
    });
    console.log("Token acquired silently");
    return response.accessToken;
  } catch (silentError) {
    console.warn("Silent token acquisition failed:", silentError);

    if (silentError instanceof InteractionRequiredAuthError) {
      // fallback to popup if silent fails
      try {
        const response = await msalInstance.acquireTokenPopup({
          scopes,
          account,
        });
        console.log("Token acquired via popup");
        return response.accessToken;
      } catch (popupError) {
        console.error("acquireTokenPopup failed:", popupError);
        return undefined;
      }
    }
    return undefined;
  }
}
