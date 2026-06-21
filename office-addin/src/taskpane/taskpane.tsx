import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import "./styles/taskpane.css";

// Mount React immediately so UI is visible even if Office.js is slow to load
const container = document.getElementById("root");
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}

// Initialize Office startup behavior after Office.js is ready
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const win = window as any;
if (typeof win.Office !== "undefined" && win.Office.onReady) {
  win.Office.onReady(async () => {
    try {
      if (win.Office.addin && win.Office.addin.setStartupBehavior) {
        await win.Office.addin.setStartupBehavior(win.Office.StartupBehavior.load);
      }
    } catch {
      // Older Office versions may not support setStartupBehavior — safe to ignore.
    }
  });
} else {
  // Office.js not yet available — poll until it is (handles async script load)
  const poll = setInterval(() => {
    if (typeof win.Office !== "undefined" && win.Office.onReady) {
      clearInterval(poll);
      win.Office.onReady(async () => {
        try {
          if (win.Office.addin?.setStartupBehavior) {
            await win.Office.addin.setStartupBehavior(win.Office.StartupBehavior.load);
          }
        } catch { /* ignore */ }
      });
    }
  }, 200);
  // Give up polling after 30s
  setTimeout(() => clearInterval(poll), 30000);
}
