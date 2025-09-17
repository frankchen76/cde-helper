import * as restify from "restify";
import { ServerSideAuthService } from "../services/auth/ServerSideAuthService";
import config from "../config";
import { err, info } from '../services/log';

export const TokenRoute = (server) => {
    server.post('/api/refreshtoken', async (req, res) => {
        const authService = new ServerSideAuthService(config.azureDevOpsProviderConfig);
        let token;
        let errorMessage = "";
        try {
            token = await authService.refreshToken(req.body.refreshToken, req.body.scopes);
            res.send(200, token);
        } catch (error) {
            err(error);
            res.send(400, {
                error: err
            });
        }
    });
    server.post('/api/gettokenbyauthcode', async (req, res) => {
        const authService = new ServerSideAuthService(config.azureDevOpsProviderConfig);
        let token;
        let errorMessage = "";
        try {
            info(`gettokenbyauthcode authcode: ${req.body.code}; host: ${req.header('Host')}; scopes: ${req.body.scopes}`);
            token = await authService.getTokenByCode(req.body.code, req.body.scopes);
            info(`gettokenbyauthcode-token:`, token);
            res.send(200, token);
        } catch (err) {
            err("/api/gettokenbyauthcode", err);
            res.send(400, {
                error: err
            });
        }
    });

}