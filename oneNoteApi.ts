import { PublicClientApplication } from "@azure/msal-browser";
import { Client, AuthenticationProvider } from "@microsoft/microsoft-graph-client";

const msalConfig = {
  auth: {
    clientId: "",
    authority: "",
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

// Acquire an access token
const loginRequest = {
  scopes: ["openid", "profile", "Notes.Read", "Notes.ReadWrite"], // Add any additional scopes as needed
};

const authResult = await msalInstance.loginPopup(loginRequest);

console.log(authResult)

// Create a Graph client instance
