import App from "./components/1.A.App";
import * as excel from "./components/Excel.App";
import * as onenote from "./components/OneNote.App";
import * as outlook from "./components/Outlook.App";
import * as powerpoint from "./components/PowerPoint.App";
import * as project from "./components/Project.App";
import * as word from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";


/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Nietiet GmbH Office Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
