// Import required packages
import * as restify from "restify";
import { info } from './services/log';
import { TasksReportRoute } from "./routers/TasksReportRoute";
import { SettingsRoute } from "./routers/SettingsRoute";
import { TokenRoute } from "./routers/TokenRoute";
import { InitPassport } from "./services/auth/PassportHelper";
import { BotRoute } from "./routers/BotRoute";
import { PageRoute } from "./routers/PageRoute";

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.queryParser());
server.use(restify.plugins.bodyParser());
server.use(restify.plugins.gzipResponse());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    info(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// add bot route
BotRoute(server);

const passport = InitPassport(server);

// add setting route
SettingsRoute(server, passport);

// add token route
TokenRoute(server);

// add tasks report route
TasksReportRoute(server, passport);

// Serve a static web page
PageRoute(server);