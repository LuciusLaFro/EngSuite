import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import ExcelTaskPane from "./ExcelTaskPane";
import { LoadSettings } from "../GlobalDeclarations";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

async function render(Component){
  let setter = await LoadSettings();
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component isOfficeInitialized={isOfficeInitialized} settings={setter}/>
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  
  render(ExcelTaskPane);
});

if ((module as any).hot) {
  (module as any).hot.accept("./ExcelTaskPane", () => {
    const NextApp = require("./ExcelTaskPane").default;
    render(NextApp);
  });
}
