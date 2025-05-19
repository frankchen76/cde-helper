import { authentication, app } from "@microsoft/teams-js";
import { IToken } from "../../services/auth/IToken";
import { IAuthCode } from "../../services/auth/IAuthCode";
import { AuthServiceToken, AuthServiceTokenCache } from "./AuthServiceToken";
import { error, info } from "./log";
import { HostInfo, OriginType } from "./HostInfo";
import { msalInstance } from "../taskpane";
import _ from "lodash";


export enum ScopesEnum {
    AzureDevOps = "https://app.vssps.visualstudio.com/.default offline_access",
    CustomApi = "api://e46dac3e-acfb-48a2-a65e-8874c0b98ee3/.default offline_access",
};

var _IAuthService: IAuthService = null;
export const getAuthService = () => {
    const hostInfo = HostInfo.getHostInfo();
    if (_IAuthService == null) {
        // check if we are in Outlook taskpane or web app
        switch (hostInfo.Origin) {
            case OriginType.OutlookTaskPane:
                _IAuthService = new OutlookAuthService();
                break;
            case OriginType.MSTeams:
                _IAuthService = new MSTeamsAuthService();
                break;
            default:
                _IAuthService = new MsalAuthService();
                break;
        }
    }
    return _IAuthService;
}

export interface IAuthService {
    getAccessToken(scopes: ScopesEnum): Promise<IToken>;
}

export abstract class ClientSideAuthServiceBase implements IAuthService {
    protected static TENANT_ID = "6f423eb7-7932-4e19-ae14-fa375038681b";
    protected static CLIENT_ID = "d5be9481-3999-4101-b0a2-99834cf4c1ad";

    protected abstract getIToken(scopes: ScopesEnum): Promise<IToken>;
    public async getAccessToken(scopes: ScopesEnum): Promise<IToken> {
        let tokenCache = AuthServiceTokenCache.createInstanceFromCache();
        let existToken: AuthServiceToken = tokenCache.getToken(scopes);

        if (existToken && !existToken.IsAccessTokenValid) {
            //refresh token based on refresh_token if token expired
            try {
                info(`Token expired, refresh token...`, existToken);
                const newToken = await this.refreshIToken(existToken.refresh_token, scopes);
                if (newToken) {
                    existToken = AuthServiceToken.createInstanceFromIToken(newToken, scopes);
                    // save refreshed token
                    tokenCache.addToken(existToken);
                    tokenCache.saveTokenCache();
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

        // kickoff authentication if token not exist or expired
        if (existToken == null) {
            info(`Retrieve token based on prompt.`);
            const newToken = await this.getIToken(scopes);
            existToken = AuthServiceToken.createInstanceFromIToken(newToken, scopes);
            tokenCache.addToken(existToken);
            tokenCache.saveTokenCache();
            info(`Saved new token.`);
        }
        info(existToken);
        return existToken;
    }
    protected async getAccessTokenByCode(authCode: IAuthCode, scopes: ScopesEnum = ScopesEnum.AzureDevOps): Promise<IToken> {
        let url = `${location.protocol}//${location.host}/api/gettokenbyauthcode`;
        const tokenRequest: any = {
            code: authCode.code,
            state: authCode.state,
            scopes: scopes
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
    protected async refreshIToken(refreshToken: string, scopes: ScopesEnum = ScopesEnum.AzureDevOps): Promise<IToken> {
        let url = `${location.protocol}//${location.host}/api/refreshtoken`;
        const tokenRequest: any = {
            refreshToken: refreshToken,
            scopes: scopes
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

export class MsalAuthService extends ClientSideAuthServiceBase {

    public async getIToken(scopes: ScopesEnum = ScopesEnum.AzureDevOps): Promise<IToken> {
        let ret: IToken = null;
        //const { instance, accounts, inProgress } = useMsal();
        const request = {
            scopes: scopes.split(" "),
        }
        const result = await msalInstance.acquireTokenPopup(request)
        if (result.account) {
            msalInstance.setActiveAccount(result.account);
            ret = {
                access_token: result.accessToken,
                token_type: result.tokenType,
                refresh_token: null,
                expires_in: result.expiresOn.getTime() - new Date().getTime()
            }
        }
        return ret;
    }
}
export class MSTeamsAuthService extends ClientSideAuthServiceBase {
    protected async getIToken(scopes: ScopesEnum): Promise<IToken> {
        let ret: IToken = null;
        //const { instance, accounts, inProgress } = useMsal();
        const request = {
            scopes: scopes.split(" "),
        }
        await app.initialize();
        const context = await app.getContext();
        info("app.context:", context);
        const loginHint = "tachen@microsoft.com";
        let url = `${location.protocol}//${location.host}/web/auth-start.html?clientId=${ClientSideAuthServiceBase.CLIENT_ID}&tenantId=${ClientSideAuthServiceBase.TENANT_ID}&scope=${scopes}&loginHint=${loginHint}&stamp=${new Date().getTime()}`;
        info("url", url);
        const result = await authentication.authenticate({
            url: url,
            width: 600,
            height: 800
        })
        const authCode = JSON.parse(result) as IAuthCode;
        info("result", result);
        ret = await this.getAccessTokenByCode(authCode, scopes);
        //const callbackToken = CallbackToken.createInstance(arg.message);
        info("ret", ret);
        return ret;
    }

}

export class OutlookAuthService extends ClientSideAuthServiceBase {
    protected getIToken(scopes: ScopesEnum): Promise<IToken> {
        return new Promise<IToken>((resolve, reject) => {
            //let url = `${location.protocol}//${location.host}/login.html`;
            const loginHint = "tachen@microsoft.com";
            let url = `${location.protocol}//${location.host}/web/auth-start.html?clientId=${ClientSideAuthServiceBase.CLIENT_ID}&tenantId=${ClientSideAuthServiceBase.TENANT_ID}&scope=${scopes}&loginHint=${loginHint}&stamp=${new Date().getTime()}`;
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
                        const accessToken = this.getAccessTokenByCode(authCode, scopes);
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
    // public async getAccessToken(scopes: ScopesEnum = ScopesEnum.AzureDevOps): Promise<IToken> {
    //     let tokenCache = AuthServiceTokenCache.createInstanceFromCache();
    //     let existToken: AuthServiceToken = tokenCache.getToken(scopes);

    //     if (existToken && !existToken.IsAccessTokenValid) {
    //         //refresh token based on refresh_token if token expired
    //         try {
    //             info(`Token expired, refresh token...`, existToken);
    //             const newToken = await this.refreshAccessToken(existToken.refresh_token, scopes);
    //             if (newToken) {
    //                 existToken = AuthServiceToken.createInstanceFromIToken(newToken, scopes);
    //                 // save refreshed token
    //                 tokenCache.addToken(existToken);
    //                 tokenCache.saveTokenCache();
    //                 info(`Saved refreshed token.`);
    //             } else {
    //                 existToken = null;
    //                 info("Cannot refresh token.");
    //             }
    //         } catch (err) {
    //             existToken = null;
    //             error("Refresh token failed", err);
    //         }
    //     } else {
    //         info(`Reuse existed token.`);
    //     }

    //     // kickoff authentication if token not exist or expired
    //     if (existToken == null) {
    //         info(`Retrieve token based on prompt.`);
    //         const newToken = await this.getIToken(scopes);
    //         existToken = AuthServiceToken.createInstanceFromIToken(newToken, scopes);
    //         tokenCache.addToken(existToken);
    //         tokenCache.saveTokenCache();
    //         info(`Saved new token.`);
    //     }
    //     info(existToken);
    //     return existToken;
    // }

    // private async getAccessTokenByCode(authCode: IAuthCode, scopes: ScopesEnum = ScopesEnum.AzureDevOps): Promise<IToken> {
    //     let url = `${location.protocol}//${location.host}/api/gettokenbyauthcode`;
    //     const tokenRequest: any = {
    //         code: authCode.code,
    //         state: authCode.state,
    //         scopes: scopes
    //     };
    //     const tokenResponse = await fetch(url, {
    //         method: 'POST',
    //         headers: {
    //             'Content-Type': 'application/json',
    //             'Accept': 'application/json',
    //             'cache': "no-store"
    //         },
    //         body: JSON.stringify(tokenRequest)
    //     });
    //     if (!tokenResponse.ok) {
    //         throw await tokenResponse.text();
    //     }
    //     return tokenResponse.json();

    // }
}