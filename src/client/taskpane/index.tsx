//import "@fluentui/react/dist/css/fabric.min.css";
import App from "../components/App";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import { createRoot } from 'react-dom/client';
import { PartialTheme, Theme, ThemeProvider } from "@fluentui/react";
/* global AppCpntainer, Component, document, Office, module, require */
// import '../../public/styles/taskpane.css';
initializeIcons();

let isOfficeInitialized = false;
let currentOfficeTheme: Office.OfficeTheme;

const currentUrl = new URL(window.location.href);
const origin = currentUrl.searchParams.get("origin");

const title = "Contoso Task Pane Add-in";

const render = () => {
    const root = createRoot(
        document.getElementById('container') as HTMLElement
    );
    root.render(
        <ThemeProvider >
            <App title={title} isOfficeInitialized={isOfficeInitialized} />
            {/* <App1 /> */}
        </ThemeProvider>,
    );
};

if (origin == "msteams") {
    render();
} else {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Outlook) {
            // document.getElementById("sideload-msg").style.display = "none";
            // document.getElementById("app-body").style.display = "flex";
            // document.getElementById("run").onclick = run;
            render();
        }
    });
}



