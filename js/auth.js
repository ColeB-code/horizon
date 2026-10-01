// MSAL configuration and authentication helpers for Family Horizon.

const msalConfig = {
  auth: {
    clientId: "4b498d6c-6c07-495d-b6fc-c31b001957ef",
    authority: "https://login.microsoftonline.com/consumers",
    redirectUri: window.location.hostname === "localhost"
      ? window.location.origin
      : `${window.location.origin}${window.location.pathname}`
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const SCOPES = ["Calendars.Read"];

let msalInstance = null;

export async function initAuth() {
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();
  return msalInstance;
}

export function getActiveAccount() {
  const accounts = msalInstance.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

export async function signIn() {
  const result = await msalInstance.loginPopup({ scopes: SCOPES });
  msalInstance.setActiveAccount(result.account);
  return result.account;
}

export async function signOut() {
  const account = getActiveAccount();
  await msalInstance.logoutPopup({ account });
}

export async function getAccessToken() {
  const account = getActiveAccount();
  if (!account) {
    throw new Error("No signed-in account.");
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: SCOPES,
      account
    });
    return result.accessToken;
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
      return result.accessToken;
    }
    throw error;
  }
}
