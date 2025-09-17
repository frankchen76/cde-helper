import * as restify from "restify";
const path = require("path");

export const PageRoute = (server) => {
    server.get(
        //"/auth-:name(start|end|config).html",
        "/web/*",
        restify.plugins.serveStatic({
            //directory: path.join(__dirname, "public"),
            directory: path.join(__dirname),
            gzip: true,
        })
    );

}