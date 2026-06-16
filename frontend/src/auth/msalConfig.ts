import { PublicClientApplication, type Configuration } from "@azure/msal-browser";

const clientId = import.meta.env.VITE_AZURE_CLIENT_ID as string | undefined;
const tenantId = import.meta.env.VITE_AZURE_TENANT_ID as string | undefined;

export const isMsalConfigured = !!(clientId && tenantId);

const config: Configuration = {
  auth: {
    clientId: clientId ?? "not-configured",
    authority: `https://login.microsoftonline.com/${tenantId ?? "common"}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const msalInstance = new PublicClientApplication(config);

export const APP_SCOPES = ["User.Read"];
export const PLANNER_SCOPES = ["Tasks.ReadWrite"];
export const CALENDAR_SCOPES = ["Calendars.Read"];
