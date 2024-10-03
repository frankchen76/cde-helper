import { IToken } from "../../services/auth/IToken";
import { IAuthCode } from "../../services/auth/IAuthCode";
import { ICallbackModel } from "../../services/auth/ICallbackModel";
import { AuthServiceToken, CallbackToken } from "./AuthServiceToken";
import { error, info } from "./log";
import { v4 as uuidv4 } from 'uuid';


export class ClientSideAuthService {
    private static CACHE_TOKEN = "cde-helper-token";
    //private static _token: IToken = null; 
    private static TENANT_ID = "6f423eb7-7932-4e19-ae14-fa375038681b";
    private static CLIENT_ID = "d5be9481-3999-4101-b0a2-99834cf4c1ad";
    private static SCOPE = "https://app.vssps.visualstudio.com/.default";

    public async getAccessToken(): Promise<IToken> {
        //const existTokenJson = null;
        const existTokenJson = localStorage.getItem(ClientSideAuthService.CACHE_TOKEN);
        let existToken: AuthServiceToken = null;
        if (existTokenJson) {
            existToken = AuthServiceToken.createInstanceFromJSON(existTokenJson);
        }

        if (existToken && !existToken.IsAccessTokenValid) {
            //refresh token based on refresh_token if token expired
            try {
                info(`Token expired, refresh token...`, existToken);
                const refreshToken = await this.refreshAccessToken(existToken.refresh_token);
                if (refreshToken) {
                    existToken = AuthServiceToken.createInstanceFromIToken(refreshToken);
                    // save refreshed token
                    localStorage.setItem(ClientSideAuthService.CACHE_TOKEN, existToken.toJson());
                    info(`Saved refreshed token.`);
                } else {
                    existToken = null;
                    info("Cannot refresh token.");
                }
            } catch (err) {
                existToken = null;
                error("Refresh token failed", err);
            }
        } else {
            info(`Reuse existed token.`);
        }

        if (existToken == null) {
            info(`Retrieve token based on prompt.`);
            existToken = AuthServiceToken.createInstanceFromIToken(await this.getIToken());
            localStorage.setItem(ClientSideAuthService.CACHE_TOKEN, existToken.toJson());

        }
        info(existToken);
        return existToken;
    }

    public getIToken(): Promise<IToken> {
        return new Promise<IToken>((resolve, reject) => {
            //let url = `${location.protocol}//${location.host}/login.html`;
            const loginHint = "tachen@microsoft.com";
            let url = `${location.protocol}//${location.host}/web/auth-start.html?clientId=${ClientSideAuthService.CLIENT_ID}&tenantId=${ClientSideAuthService.TENANT_ID}&scope=${ClientSideAuthService.SCOPE}&loginHint=${loginHint}&stamp=${new Date().getTime()}`;
            //info(`open dialog ${url}`);
            let dialog;
            const w = 600 / screen.width * 100;
            const h = 800 / screen.height * 100;
            Office.context.ui.displayDialogAsync(url, { height: h, width: w }, (asyncResult: Office.AsyncResult<Office.Dialog>) => {
                if (asyncResult.status.toString() == "failed") {
                    // In addition to general system errors, there are 3 specific errors for 
                    // displayDialogAsync that you can handle individually.
                    switch (asyncResult.error.code) {
                        case 12004:
                            info("Domain is not trusted");
                            break;
                        case 12005:
                            info("HTTPS is required");
                            break;
                        case 12007:
                            info("A dialog is already opened.");
                            break;
                        default:
                            info(asyncResult.error.message);
                            break;
                    }
                }
                else {
                    dialog = asyncResult.value;
                    /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...) ff*/
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: { message: string, origin: string | undefined }) => {
                        // info("token:");
                        // info(arg);
                        // setToken(`token: ${arg.message}`);
                        const authCode = JSON.parse(arg.message) as IAuthCode;
                        const accessToken = this.getAccessTokenByCode(authCode);
                        //const callbackToken = CallbackToken.createInstance(arg.message);
                        info(accessToken);
                        resolve(accessToken);
                        dialog.close();
                    });

                    /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
                    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: { error: number, type: string }) => {
                        let errMsg = "";
                        switch (arg.error) {
                            case 12002:
                                errMsg = "Cannot load URL, no such page or bad URL syntax.";
                                break;
                            case 12003:
                                errMsg = "HTTPS is required.";
                                break;
                            case 12006:
                                // The dialog was closed, typically because the user the pressed X button.
                                errMsg = "Dialog closed by user";
                                break;
                            default:
                                errMsg = "Undefined error in dialog window";
                                break;
                        }
                        if (dialog) dialog.close();
                        reject(errMsg);
                    });
                }
            });

        });
    }
    public async getAccessTokenByCode(authCode: IAuthCode): Promise<IToken> {
        let url = `${location.protocol}//${location.host}/api/gettokenbyauthcode`;
        const tokenRequest: any = {
            code: authCode.code,
            state: authCode.state
        };
        const tokenResponse = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'cache': "no-store"
            },
            body: JSON.stringify(tokenRequest)
        });
        if (!tokenResponse.ok) {
            throw await tokenResponse.text();
        }
        return tokenResponse.json();

    }
    public async refreshAccessToken(refreshToken: string): Promise<IToken> {
        let url = `${location.protocol}//${location.host}/api/refreshtoken`;
        const tokenRequest: any = {
            refreshToken: refreshToken
        };
        const tokenResponse = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'cache': "no-store"
            },
            body: JSON.stringify(tokenRequest)
        });
        if (!tokenResponse.ok) {
            throw await tokenResponse.text();
        }
        return tokenResponse.json();
    }
}