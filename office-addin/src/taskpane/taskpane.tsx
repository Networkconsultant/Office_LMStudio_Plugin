import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import "./styles/taskpane.css";

/* Wait for Office.js to be ready before mounting React. */
Office.onReady(async () => {
  const container = document.getElementById("root");
  if (!container) return;
  const root = createRoot(container);
  root.render(<App />);

  /* Make the task pane auto-open on every subsequent document open. */
  try {
    if (Office.addin && Office.addin.setStartupBehavior) {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }
  } catch {
    // Older Office versions may not support setStartupBehavior — safe to ignore.
  }
});
