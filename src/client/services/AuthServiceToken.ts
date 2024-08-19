import { IToken } from "../../services/auth/IToken";
import moment from 'moment';
import { jwtDecode } from "jwt-decode";
import { ICallbackModel } from "../../services/auth/ICallbackModel";
import { info } from "./log";;

export class AuthServiceToken implements IToken {
    constructor(public access_token: string,
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
    public static createInstanceFromIToken(token: IToken) {
        return new AuthServiceToken(token.access_token,
            token.refresh_token,
            token.token_type,
            token.expires_in);
    }
    public static createInstanceFromJSON(json: string) {
        let ret: AuthServiceToken = null;
        try {
            const item = JSON.parse(json) as IToken;
            ret = new AuthServiceToken(item.access_token,
                item.refresh_token,
                item.token_type,
                item.expires_in);

        } catch (error) {
            info(`Cannot deserialized token object from cache. '${json}'`);
        }
        return ret;
    }
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
