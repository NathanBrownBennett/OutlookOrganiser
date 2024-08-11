import * as msal from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    redirectUri: process.env.REDIRECT_URI,
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

export async function getAccessToken() {
  const loginRequest = {
    scopes: ["User.Read", "Mail.Read", "Files.ReadWrite.All"],
  };

  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    return loginResponse.accessToken;
  } catch (error) {
    console.error("Authentication failed: ", error);
  }
}
