import * as restify from "restify";
import { err, info } from '../services/log';
import { SettingsDbSerivce } from "../services/db/SettingsDbSerivce";
const path = require("path");
const fs = require('fs');

export const SettingsRoute = (server, passport) => {
    const env = process.env.NODE_ENV || 'production';

    server.get('/api/settings', passport.authenticate('oauth-bearer', { session: false }), async (req, res) => {
        const settingsService = new SettingsDbSerivce();
        try {
            info("upn:", req.query.upn);
            const result = await settingsService.getSettings(req.query.upn);
            info("Get settings:", result);
            res.send(200, result);
        } catch (err) {
            console.error(err);
            res.send(500, {
                error: err
            });
        }
    });
    server.post('/api/settings', passport.authenticate('oauth-bearer', { session: false }), async (req: restify.Request, res: restify.Response) => {
        const settingsService = new SettingsDbSerivce();
        try {
            info("upn:", req.query.upn);
            info("body:", JSON.stringify(req.body, null, 4));
            const result = await settingsService.saveSettings(req.query.upn, req.body);
            if (result && (result.statusCode == 200 || result.statusCode == 201)) {
                info("Save settings:", result);
                res.send(200, result);
            } else {
                res.send(400, "Failed to save the settings.");
            }
        } catch (err) {
            console.error(err);
            res.send(500, {
                error: err
            });
        }
    });

}