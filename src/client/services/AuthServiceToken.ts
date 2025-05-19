import { IToken } from "../../services/auth/IToken";
import moment from 'moment';
import { jwtDecode } from "jwt-decode";
import { ICallbackModel } from "../../services/auth/ICallbackModel";
import { info } from "./log";;

export class AuthServiceTokenCache {
    private static CACHE_TOKEN = "cde-helper-token";
    private _tokens: AuthServiceToken[] = [];

    public addToken(token: AuthServiceToken) {
        // check if token already exist, remove it first
        const index = this._tokens.findIndex(t => t.scopes == token.scopes);
        if (index > -1) {
            this._tokens.splice(index, 1);
        }
        this._tokens.push(token);
    }
    public getToken(scopes: string): AuthServiceToken {
        const ret = this._tokens.find(t => t.scopes == scopes);
        return ret;
    }
    public saveTokenCache() {
        const tokenCacheJson = JSON.stringify(this);
        localStorage.setItem(AuthServiceTokenCache.CACHE_TOKEN, tokenCacheJson);
        info(`Token cache saved. ${tokenCacheJson}`);
    }
    public static createInstanceFromCache(): AuthServiceTokenCache {
        const existTokenJson = localStorage.getItem(AuthServiceTokenCache.CACHE_TOKEN);
        let existTokenCache: AuthServiceTokenCache = null;
        if (existTokenJson) {
            existTokenCache = AuthServiceTokenCache.createInstance(existTokenJson);
        } else {
            existTokenCache = new AuthServiceTokenCache();
        }
        return existTokenCache;
    }
    private static createInstance(json: string): AuthServiceTokenCache {
        let ret: AuthServiceTokenCache = null;
        try {
            const item = JSON.parse(json) as AuthServiceTokenCache;
            ret = new AuthServiceTokenCache();
            item._tokens.forEach(t => {
                const token = new AuthServiceToken(t.scopes,
                    t.access_token,
                    t.refresh_token,
                    t.token_type,
                    t.expires_in);
                ret.addToken(token);
            });

        } catch (error) {
            info(`Cannot deserialized token object from cache. '${json}'`);
        }
        return ret;
    }
}
export class AuthServiceToken implements IToken {
    constructor(public scopes: string, public access_token: string,
        public refresh_token: string,
        public token_type: string,
        public expires_in: number) {

    }
    public get IsAccessTokenValid(): boolean {
        let ret = false;
        info(this);
        if (this.access_token) {
            // 
            const decodedToken = jwtDecode(this.access_token);
            const tokenExpired = moment(new Date(decodedToken.exp * 1000));
            ret = tokenExpired > moment();
            // info("tokenExpired:");
            // info(tokenExpired);
            // info(moment());
        }
        return ret;
    }
    public static createInstanceFromIToken(token: IToken, scopes: string): AuthServiceToken {
        return new AuthServiceToken(scopes,
            token.access_token,
            token.refresh_token,
            token.token_type,
            token.expires_in);
    }
    // public static createInstanceFromJSON(json: string) {
    //     let ret: AuthServiceToken = null;
    //     try {
    //         const item = JSON.parse(json) as IToken;
    //         ret = new AuthServiceToken(item.access_token,
    //             item.refresh_token,
    //             item.token_type,
    //             item.expires_in);

    //     } catch (error) {
    //         info(`Cannot deserialized token object from cache. '${json}'`);
    //     }
    //     return ret;
    // }
    public toJson(): string {
        return JSON.stringify(this);
    }
}
export class CallbackToken implements ICallbackModel {
    // access_token: string;
    // token_type: string;
    // refresh_token: string;
    // expires_in: number;
    // error: string;

    public constructor(public access_token: string,
        public token_type: string,
        public refresh_token: string,
        public expires_in: number,
        public error: string) {

    }
    public get HasError(): boolean {
        return this.error != null && this.error != "";// || this.error==undefined || this.error=="";
    }
    public get Token(): IToken {
        return {
            access_token: this.access_token,
            token_type: this.token_type,
            refresh_token: this.refresh_token,
            expires_in: this.expires_in
        };
    }

    public static createInstance(json: string): CallbackToken {
        const result = JSON.parse(json) as CallbackToken;
        return new CallbackToken(result.access_token,
            result.token_type,
            result.refresh_token,
            result.expires_in,
            result.error);
    }
    public toJson(): string {
        return JSON.stringify(this);
    }
}
