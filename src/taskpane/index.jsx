import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require */

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      {/* FIX #21: Removed unused 'title' prop (App.jsx does not consume it) */}
      <App />
    </FluentProvider>
  );
});

/* FIX #22: Hot module replacement must render a React element, not the component class */
if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <NextApp />
      </FluentProvider>
    );
  });
}
