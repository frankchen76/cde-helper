//import "@fluentui/react/dist/css/fabric.min.css";
import App from "../components/App";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import { createRoot } from 'react-dom/client';
import { PartialTheme, Theme, ThemeProvider } from "@fluentui/react";
import { HostInfo, OriginType } from "../services/HostInfo";
// MSAL imports
import {
    PublicClientApplication,
    Configuration,
    EventType,
    EventMessage,
    AuthenticationResult,
} from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import * as microsoftTeams from "@microsoft/teams-js";

/* global AppCpntainer, Component, document, Office, module, require */
// import '../../public/styles/taskpane.css';
initializeIcons();

let currentOfficeTheme: Office.OfficeTheme;

//const currentUrl = new URL(window.location.href);
//const origin = currentUrl.searchParams.get("origin");
const hostInfo = HostInfo.getHostInfo();

const title = "Contoso Task Pane Add-in";

const root = createRoot(
    document.getElementById('container') as HTMLElement
);
export var msalInstance: PublicClientApplication;
switch (hostInfo.Origin) {
    case OriginType.MSTeams:
        microsoftTeams.app.initialize();
        root.render(
            <ThemeProvider >
                <App title={title} hostInfo={hostInfo} />
                {/* <App1 /> */}
            </ThemeProvider>,
        );
        break;
    case OriginType.OutlookTaskPane:
        root.render(
            <ThemeProvider >
                <App title={title} hostInfo={hostInfo} />
                {/* <App1 /> */}
            </ThemeProvider>,
        );
        break;
    default:
        const msalConfig: Configuration = {
            auth: {
                clientId: "d5be9481-3999-4101-b0a2-99834cf4c1ad",
                authority: "https://login.microsoftonline.com/6f423eb7-7932-4e19-ae14-fa375038681b",
                redirectUri: "/web/teampane.html",
                postLogoutRedirectUri: "/",
            },
            system: {
                allowPlatformBroker: false, // Disables WAM Broker
            },
        };
        msalInstance = new PublicClientApplication(msalConfig);
        msalInstance.initialize().then(() => {
            // Account selection logic is app dependent. Adjust as needed for different use cases.
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.setActiveAccount(accounts[0]);
            }

            msalInstance.addEventCallback((event: EventMessage) => {
                if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
                    const payload = event.payload as AuthenticationResult;
                    const account = payload.account;
                    msalInstance.setActiveAccount(account);
                }
            });

            root.render(
                <MsalProvider instance={msalInstance}>
                    <ThemeProvider >
                        <App title={title} hostInfo={hostInfo} />
                        {/* <App1 /> */}
                    </ThemeProvider>
                </MsalProvider>
            );
        });
        break;
}


// const render = () => {
//     const root = createRoot(
//         document.getElementById('container') as HTMLElement
//     );
//     root.render(
//         <ThemeProvider >
//             <App title={title} isOfficeInitialized={isOfficeInitialized} />
//             {/* <App1 /> */}
//         </ThemeProvider>,
//     );
// };

// if (origin == "msteams") {
//     render();
// } else {
//     Office.onReady((info) => {
//         if (info.host === Office.HostType.Outlook) {
//             // document.getElementById("sideload-msg").style.display = "none";
//             // document.getElementById("app-body").style.display = "flex";
//             // document.getElementById("run").onclick = run;
//             render();
//         }
//     });
// }



