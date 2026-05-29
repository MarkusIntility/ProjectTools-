import "@intility/bifrost-css/dist/bifrost-bundle.css";
import { Bifrost } from "@intility/bifrost-react";
// @ts-expect-error – locale files are not re-exported via the main types entry in v3
import nbNO from "@intility/bifrost-react/dist/locales/nb-no";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { BrowserRouter } from "react-router-dom";
import App from "./App";
import { msalInstance, isMsalConfigured } from "./auth/msalConfig";

async function bootstrap() {
  if (isMsalConfigured) {
    try {
      // Must run before React renders so ?code= in URL is read before
      // React Router's <Navigate> strips it from window.location.
      await msalInstance.initialize();

      // MSAL v3 does not auto-navigate to redirectStartPage after loginRedirect.
      // We store the return URL ourselves and navigate here after successful auth.
      const returnUrl = sessionStorage.getItem("app.returnUrl");
      if (returnUrl) {
        sessionStorage.removeItem("app.returnUrl");
        // If auth succeeded there will be at least one account in the cache
        if (msalInstance.getAllAccounts().length > 0) {
          window.location.replace(returnUrl);
          return; // Don't render React — a new page load is starting
        }
        // Auth failed/cancelled — fall through and render at current location
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
