import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require, HTMLElement, window */

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
if (!rootElement) {
  console.error("Container element not found!");
}

const root = rootElement ? createRoot(rootElement) : undefined;

const renderApp = () => {
  if (!root) {
    console.error("React root not initialized!");
    return;
  }
  try {
    root.render(
      <FluentProvider theme={webLightTheme}>
        <App title={title} />
      </FluentProvider>
    );
  } catch (error) {
    console.error("Error rendering app:", error);
  }
};

/* Render application after Office initializes */
const initializeApp = () => {
  console.log("Initializing Office add-in...");
  // Standard Office.js initialization pattern
  let attempts = 0;
  const maxAttempts = 60; // 3 seconds max (60 * 50ms)
  let appRendered = false;
  
  const waitForOffice = (): void => {
    if (appRendered) return; // Prevent multiple renders
    
    attempts++;
    
    // Check if Office is available (could be on window or global)
    const OfficeObj = (typeof Office !== "undefined" ? Office : (window as any).Office) as any;
    
    if (OfficeObj && typeof OfficeObj.onReady === "function") {
      // Office.js is loaded, use onReady
      appRendered = true;
      OfficeObj.onReady()
        .then((info: any) => {
          console.log("Office initialized successfully", info);
          renderApp();
        })
        .catch((error: any) => {
          console.error("Office.onReady failed:", error);
          // Still render - might work in browser or Office might be partially available
          renderApp();
        });
    } else if (attempts < maxAttempts) {
      // Office.js not loaded yet, wait a bit and retry
      setTimeout(() => {
        waitForOffice();
      }, 50);
    } else {
      // Max attempts reached, render anyway (browser mode or Office.js failed to load)
      if (!appRendered) {
        appRendered = true;
        console.log("Office.js not detected after waiting, rendering in browser mode");
        renderApp();
      }
    }
  };
  
  // Start waiting for Office.js
  waitForOffice();
};

// Wait for DOM to be ready, then initialize
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initializeApp);
} else {
  // DOM is already ready
  initializeApp();
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    if (root) {
      root.render(
        <FluentProvider theme={webLightTheme}>
          <NextApp title={title} />
        </FluentProvider>
      );
    }
  });
}
