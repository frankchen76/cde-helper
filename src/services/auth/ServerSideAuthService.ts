import { IToken } from "./IToken";
import { Request } from "restify";
import * as https from "https";
import { AzureDevOpsTokenProvider, IAzureDevOpsProviderConfig } from "../AzureDevOpsTokenProvider";

export enum GrantTypeEnum {
    Callback = "urn:ietf:params:oauth:grant-type:jwt-bearer",
    RefreshToken = "refresh_token"
}
export class ServerSideAuthService {
    constructor(private config: IAzureDevOpsProviderConfig) {
    }
    public async refreshToken(refreshToken: string): Promise<IToken> {
        //return await this.retrieveToken(req, GrantTypeEnum.RefreshToken, refreshToken);
        const token = await AzureDevOpsTokenProvider.refreshAccessToken(this.config, refreshToken);
        return token;
    }
    public async getTokenByCode(authCode: string): Promise<IToken> {
        const token = await AzureDevOpsTokenProvider.getAccessTokenByCode(this.config, authCode);
        return token;
    }
}