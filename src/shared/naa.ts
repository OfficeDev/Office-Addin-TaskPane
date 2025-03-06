/**
 * @file : naa.js
 * @author : Microsoft Corporation
 */

import { 
  AuthenticationResult, 
  createNestablePublicClientApplication, 
  IPublicClientApplication 
} from "@azure/msal-browser";

/* global console Office */

const applicationId = "Enter_the_Application_Id_Here";

var naaPca: IPublicClientApplication;

async function initializePca(clientId: string): Promise<void> {
  if (naaPca == null) {
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: "https://login.microsoftonline.com/common",
        supportsNestedAppAuth: true,
      },
    };

    naaPca = await createNestablePublicClientApplication(msalConfig);
  }
}

export async function getAuthentication(scopes: string[]): Promise<AuthenticationResult> {
  var clientId;

  if (Office.context.urls && Office.context.urls.javascriptRuntimeUrl) {
    const matchResult = Office.context.urls.javascriptRuntimeUrl.match(/clientid=([^&]+)/);
    if (!matchResult) {
      throw new Error("Client ID not found in the URL");
    }
    clientId = matchResult[1];
  } else {
    // fallback to prod client ID. Mobile does not support javascriptRuntimeUrl API yet
    clientId = applicationId;
  }

  let authentication: AuthenticationResult;
  const tokenRequest = {
    scopes: scopes,
    loginHint: Office.context.mailbox.userProfile.emailAddress,
  };

  await initializePca(clientId);
  try {
    authentication = await naaPca.acquireTokenSilent(tokenRequest);
  } catch (error) {
    console.log("Failed to get token silently, re-trying with popup.");
    console.log(error);

    // Acquire token silent failure. Send an interactive request via popup.
    authentication = await naaPca.acquireTokenPopup(tokenRequest);

    if (authentication == null) {
      throw "Account is null - token popup was canceled by user?";
    }
  }

  return authentication;
}
