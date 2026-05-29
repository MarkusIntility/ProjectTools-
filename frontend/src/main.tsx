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
    // Initialize MSAL before React renders so handleRedirectPromise() processes
    // the ?code= from loginRedirect before React Router's <Navigate> strips them.
    try {
      await msalInstance.initialize();
      // After initialize(), MSAL auto-navigates to redirectStartPage if returning
      // from loginRedirect. If that navigation fires, this code won't continue.
    } catch (err) {
      // Store error so RunbookPage can display it after redirect
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
