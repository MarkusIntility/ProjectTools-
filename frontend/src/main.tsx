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
    // Initialize MSAL BEFORE React renders so URL params (?code=, ?state=) are
    // read before React Router's <Navigate> strips them from window.location.
    await msalInstance.initialize();

    // If this window is an MSAL popup (opened by loginPopup), MSAL has already
    // processed the auth response above. Skip rendering React — the popup will
    // close itself automatically.
    if (window.opener !== null) return;
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
