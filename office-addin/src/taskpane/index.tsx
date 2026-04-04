/**
 * index.tsx – React-Einstiegspunkt für die Infront Toolkit Task Pane.
 * Wartet auf Office.onReady() bevor React gemountet wird.
 */

import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./App";

/* global Office */

Office.onReady(() => {
  ReactDOM.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
    document.getElementById("root")
  );
});
