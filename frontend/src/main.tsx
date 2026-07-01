import "@intility/bifrost-css/dist/bifrost-bundle.css";
import { Bifrost } from "@intility/bifrost-react";
// @ts-expect-error – locale files are not re-exported via the main types entry in v3
import nbNO from "@intility/bifrost-react/dist/locales/nb-no";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { BrowserRouter } from "react-router-dom";
import App from "./App";
import { msalInstance, isMsalConfigured, APP_SCOPES, PLANNER_SCOPES } from "./auth/msalConfig";

async function bootstrap() {
  if (isMsalConfigured) {
    try {
      // Must initialize before any MSAL call
      await msalInstance.initialize();

      // In MSAL v5, initialize() does NOT process the redirect response.
      // We must call handleRedirectPromise() explicitly.
      // If returning from loginRedirect, this processes the ?code= from the URL
      // and stores the account in MSAL's cache.
      const redirectResult = await msalInstance.handleRedirectPromise();

      const returnUrl = sessionStorage.getItem("app.returnUrl");
      if (returnUrl) {
        sessionStorage.removeItem("app.returnUrl");
        // redirectResult is non-null when we just completed a loginRedirect
        if (redirectResult !== null) {
          window.location.replace(returnUrl);
          return; // Don't render React — a new page load is starting
        }
        // Auth failed, was cancelled, or returnUrl is stale — fall through
      }

      // Require Entra ID login — redirect if no account in cache
      if (msalInstance.getAllAccounts().length === 0) {
        sessionStorage.setItem("app.returnUrl", window.location.href);
        await msalInstance.loginRedirect({ scopes: [...APP_SCOPES, ...PLANNER_SCOPES] });
        return;
      }
    } catch (err) {
      sessionStorage.removeItem("app.returnUrl");
      sessionStorage.setItem("msal.bootError", err instanceof Error ? err.message : String(err));
    }
  }

  createRoot(document.getElementById("root")!).render(
    <StrictMode>
      <BrowserRouter>
        <Bifrost locale={nbNO}>
          <App />
        </Bifrost>
      </BrowserRouter>
    </StrictMode>
  );
}

bootstrap();
