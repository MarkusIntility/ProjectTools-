import "@intility/bifrost-css/dist/bifrost-bundle.css";
import { Bifrost } from "@intility/bifrost-react";
// @ts-expect-error – locale files are not re-exported via the main types entry in v3
import nbNO from "@intility/bifrost-react/dist/locales/nb-no";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { BrowserRouter } from "react-router-dom";
import App from "./App";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <BrowserRouter>
      <Bifrost locale={nbNO}>
        <App />
      </Bifrost>
    </BrowserRouter>
  </StrictMode>
);
