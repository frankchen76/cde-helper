// import { Request, Response } from 'express';
import { Response, Next, Request } from 'restify'
import { APIKeyDbService } from '../db/APIKeyDbSerivce';
import { IAuthCode, ITokenProvider, IUserToken, UserToken } from '../AzureDevOpsTokenProvider';
import { HttpClientService } from '../HttpClientService';
import { info, err } from '../log';

export const authenticateApiKey = async (req: Request, res: Response, next: any) => {
    const key = req.headers['x-api-key'];
    info(`Key: ${key};`);
    if (key != null) {
        const keyService = new APIKeyDbService();
        const keyItem = await keyService.authenticateKey(key.toString());
        if (keyItem) {
            // set the upn in the request
            req.params.upn = keyItem.upn;
            next();
        } else {
            res.send(401, 'Unauthorized');
        }
    } else {
        res.send(401, 'Unauthorized');
    }
}
export const authenticateBearKey = (req: Request, res: Response, next: Next) => {
    const authHeader = req.headers['authorization'];
    info('validation');
    if (authHeader) {
        info('authHeader', authHeader);
        const token = authHeader.split(' ')[1];
        const url = "https://app.vssps.visualstudio.com/_apis/profile/profiles/me?api-version=7.1-preview.3";
        const tokenProvider = new BearTokenProvider(token);
        const httpCliet = new HttpClientService(tokenProvider);
        info('validate token...', token);
        httpCliet.get(url).then((response) => {
            info('authenticateBearKey-response', response);
            if (response && response["emailAddress"]) {
                const email = response["emailAddress"];
                const authService = new APIKeyDbService();
                authService.authenticateEmail(email).then((result) => {
                    req.params.upn = response["emailAddress"];
                    return next();
                }).catch((error) => {
                    err('authenticateBearKey-emailAddress check failed', error);
                    res.send(401, 'Unauthorized');
                });
            } else {
                err('authenticateBearKey-no email');
                res.send(401, 'Unauthorized');
            }
        }).catch((error) => {
            err('authenticateBearKey-devopsprofilecheck-error', error);
            res.send(401, 'Unauthorized');
        });

    } else {
        err('authenticateBearKey-no authorization header');
        res.send(401, 'Unauthorized');
    }
}
class BearTokenProvider implements ITokenProvider {
    private _userToken: IUserToken | undefined;
    constructor(private token: string) {
        this._userToken = UserToken.createInstanceFromAccessToken(token);
    }
    getUserToken(): Promise<IUserToken> {
        throw new Error('Method not implemented.');
    }
    getAccessToken(scopes?: string[] | undefined): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            resolve(this._userToken?.accessToken!);
        });
    }
    initUserWithAuthCode(authCode: IAuthCode): Promise<void> {
        throw new Error("Method not implemented.");
    }

}