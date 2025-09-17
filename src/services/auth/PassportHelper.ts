import { err, info } from '../../services/log';
import config from "../../config";

export const InitPassport = (server): any => {
    // add the authenticateApiKey middleware to the router
    let passport = require('passport');
    let OIDCBearerStrategy = require('passport-azure-ad').BearerStrategy;
    server.use(passport.initialize()); // Starts passport
    //server.use(passport.session()); // Provides session support

    let bearerStrategy = new OIDCBearerStrategy(config.entraIdAppConfig, // config file
        // function (req, token, done) {
        //     console.log('token was the token retreived');
        //     return done(null, token);
        // }

        // passRequireCallback is set to true, so we can get the request object in the callback, otherwise you can use token and done only
        function (req, token, done) {
            //console.log(token, 'was the token retreived');
            info(token, 'was the token retreived');
            if (!token.email) {
                done(new Error('upn is not found in token'));
            }
            else {
                // add the upn to the request object for later use
                req.query.upn = token.email;
                info('creator was set to upn:', token.email);
                done(null, token);
            }
        }
    );

    passport.use(bearerStrategy);
    return passport;
}