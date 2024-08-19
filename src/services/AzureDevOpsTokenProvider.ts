import * as https from "https";
import axios, { AxiosError, AxiosResponse, type AxiosRequestConfig } from 'axios'
import { err, info } from "./log";
import { jwtDecode, JwtPayload } from "jwt-decode";
import moment from "moment";
import { StatePropertyAccessor, TurnContext } from "botbuilder";
import { v4 as uuidv4 } from 'uuid';

interface IToken {
    access_token: string;
    token_type: string;
    refresh_token: string;
    expires_in: number;
}
export interface IAzureDevOpsProviderConfig {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    projectUrl: string;
    loginUrl: string;
    redirectUrl: string;
    scopes: string | string[];
}
export interface IAuthCode {
    state: string;
    code: string;
};
export interface ITokenStore {
    saveToken(token: IUserToken): Promise<void>;
    getToken(): Promise<IUserToken>;
}

export interface ITokenProvider {
    getAccessToken(scopes?: string[]): Promise<string>;
    getUserToken(): Promise<IUserToken>;
    initUserWithAuthCode(authCode: IAuthCode): Promise<void>
}
export interface IUserToken {
    upn: string;
    verifierCode: string;
    state?: string;
    accessToken?: string;
    refreshToken?: string;
    get IsAccessTokenExpired(): boolean;
    get IsTokenValid(): boolean;
    get HasToken(): boolean;
}
export class UserToken implements IUserToken {
    private constructor(public upn: string,
        public verifierCode: string,
        public state?: string,
        public accessToken?: string,
        public refreshToken?: string) {

    }
    public get IsAccessTokenExpired(): boolean {
        let ret = false;
        if (this.accessToken) {
            const decodedToken = jwtDecode<JwtPayload>(this.accessToken);
            const tokenExpired = moment(new Date(decodedToken.exp * 1000));
            ret = tokenExpired.isBefore(moment());

        }
        return ret;
    }
    public get IsTokenValid(): boolean {
        let ret = false;
        if (this.accessToken) {
            const decodedToken = jwtDecode<JwtPayload>(this.accessToken);
            const tokenExpired = moment(new Date(decodedToken.exp * 1000));
            ret = tokenExpired.isAfter(moment());
        }
        return ret;
    }
    public get HasToken(): boolean {
        return this.accessToken !== undefined || this.refreshToken !== undefined;
    }
    public static createInstanceFromIUserToken(iUser: IUserToken): IUserToken {
        return new UserToken(iUser.upn, iUser.verifierCode, iUser.state, iUser.accessToken, iUser.refreshToken);
    }
    public static createInstanceFromVerifierCode(upn: string, verifierCode: string, state: string): IUserToken {
        return new UserToken(upn, verifierCode, state);
    }
    public static createInstanceFromAccessToken(accessToken: string): IUserToken {
        return new UserToken("", "", "", accessToken);
    }
}
export class BotStateTokenStore implements ITokenStore {
    constructor(private accessor: StatePropertyAccessor<IUserToken>, private context: TurnContext) {

    }
    async saveToken(token: IUserToken): Promise<void> {
        await this.accessor.set(this.context, token);
    }
    async getToken(): Promise<IUserToken> {
        let devOpsUser = await this.accessor.get(this.context);
        if (!devOpsUser) {
            const verifierCode = uuidv4();
            const state = uuidv4();
            devOpsUser = UserToken.createInstanceFromVerifierCode(this.context.activity.from.aadObjectId, verifierCode.toString(), state.toString());
            await this.accessor.set(this.context, devOpsUser);
            info(`Created new userToken, UPN: ${devOpsUser.upn}; verifierCode: ${devOpsUser.verifierCode}; state: ${devOpsUser.state};}`);
        } else {
            // need to reinitialize the object to have object instance. 
            devOpsUser = UserToken.createInstanceFromIUserToken(devOpsUser);
            info(`Got userToken, UPN: ${devOpsUser.upn}; verifierCode: ${devOpsUser.verifierCode}; state: ${devOpsUser.state};}`);
        }
        //info(`UserToken, UPN: ${devOpsUser.upn}; verifierCode: ${devOpsUser.verifierCode}; state: ${devOpsUser.state};}`);
        return devOpsUser;
    }

}
export class AzureDevOpsTokenProvider implements ITokenProvider {
    private userToken: IUserToken;
    private constructor(private readonly tokenStore: BotStateTokenStore, private readonly config: IAzureDevOpsProviderConfig) {

    }
    public async init(): Promise<void> {
        if (!this.userToken) {
            this.userToken = await this.tokenStore.getToken();
        }
    }
    public async getUserToken(): Promise<IUserToken> {
        return await this.tokenStore.getToken();
    }
    public async getAccessToken(scopes?: string[]): Promise<string> {
        if (!this.userToken) {
            throw new Error("userToken is not initialized");
        } else {
            if (!this.userToken.IsTokenValid) {
                const iToken = await AzureDevOpsTokenProvider.refreshAccessToken(this.config, this.userToken.refreshToken);
                this.userToken.accessToken = iToken.access_token;
                this.userToken.refreshToken = iToken.refresh_token;
            }
            await this.tokenStore.saveToken(this.userToken);
            return this.userToken.accessToken;
        }
    }
    public async initUserWithAuthCode(authCode: IAuthCode): Promise<void> {
        const iToken = await AzureDevOpsTokenProvider.getAccessTokenByCode(this.config, authCode.code);
        this.userToken.accessToken = iToken.access_token;
        this.userToken.refreshToken = iToken.refresh_token;

        // save token back to token store
        await this.tokenStore.saveToken(this.userToken);
    }

    public static async getAccessTokenByCode(config: IAzureDevOpsProviderConfig, code: string): Promise<IToken> {

        const { clientId, clientSecret, tenantId, scopes, redirectUrl } = config;
        const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

        const data = new URLSearchParams();
        data.append("client_id", clientId);
        data.append("scope", scopes.toString());
        data.append("grant_type", "authorization_code"); //"refresh_token",
        data.append("code", code); //req.body.refreshToken,
        data.append("redirect_uri", redirectUrl);
        data.append("client_secret", clientSecret);
        //data.append("code_verifier", this.userToken.verifierCode); // the verifier code is the original code without SHA256 hash

        const options: AxiosRequestConfig = {
            headers: { "Content-Type": "application/x-www-form-urlencoded" }
        };
        info("getAccessTokenByCode-url:", url);
        info("getAccessTokenByCode-data:", data.toString());
        let res: AxiosResponse<any, any> = null;
        try {
            res = await axios.post(url, data, options);
            return {
                access_token: res.data.access_token,
                refresh_token: res.data.refresh_token,
                token_type: res.data.token_type,
                expires_in: res.data.expires_in
            };

        } catch (error) {
            //err(`url: ${res.config.url}; data: ${res.config.data}; response status: ${res.status};`);
            //err("response data: ", res.data);
            err("getAccessTokenByCode-error:", error);
        }

    }
    public static async refreshAccessToken(config: IAzureDevOpsProviderConfig, refreshToken: string): Promise<IToken> {
        const { clientId, clientSecret, tenantId, scopes, redirectUrl } = config;

        let url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        const data = new URLSearchParams();
        data.append("client_id", clientId);
        data.append("scope", scopes.toString());
        //data.append("scope", "https://manage.office.com//.default");
        data.append("grant_type", "refresh_token"); //"refresh_token",
        data.append("refresh_token", refreshToken); //req.body.refreshToken,
        const options: AxiosRequestConfig = {
            headers: { "Content-Type": "application/x-www-form-urlencoded" }
        };

        const res = await axios.post(url, data, options);
        return {
            access_token: res.data.access_token,
            refresh_token: res.data.refresh_token,
            token_type: res.data.token_type,
            expires_in: res.data.expires_in
        };

    }

    public static async createInstance(tokenStore: BotStateTokenStore, config: IAzureDevOpsProviderConfig): Promise<AzureDevOpsTokenProvider> {
        const newProvider = new AzureDevOpsTokenProvider(tokenStore, config);
        await newProvider.init();
        return newProvider;
    }
}