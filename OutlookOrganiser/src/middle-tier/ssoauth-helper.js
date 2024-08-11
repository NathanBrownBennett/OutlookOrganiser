import * as msal from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID, // CLIENT_ID should be set in your .ENV file
    authority: "https://login.microsoftonline.com/common",
    redirectUri: process.env.REDIRECT_URI // REDIRECT_URI should be set in your .ENV file
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

export async function login() {
  const loginRequest = {
    scopes: ["Mail.Read", "Files.ReadWrite.All", "User.Read"]
  };

  try {
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    return loginResponse.accessToken;
  } catch (error) {
    console.error("Login failed: ", error);
  }
}
