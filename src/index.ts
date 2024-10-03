// Import required packages
import * as restify from "restify";
import { Response, Next, Request } from 'restify'

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
    CloudAdapter,
    ConfigurationServiceClientCredentialFactory,
    ConfigurationBotFrameworkAuthentication,
    TurnContext,
    MemoryStorage,
    ConversationState,
    UserState,
} from "botbuilder";
const fs = require('fs');

// This bot's main dialog.
import config from "./config";
const path = require("path");
import { TeamsBot } from "./bot/teamsBot";
import { err, info } from './services/log';
import { ServerSideAuthService } from "./services/auth/ServerSideAuthService";
import { authenticateApiKey, authenticateBearKey } from "./services/auth/APIKeyAuth";
import { CompletedTasksDbSerivce } from "./services/db/CompletedTasksDbSerivce";
import _ from "lodash";


// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: config.botId,
    MicrosoftAppPassword: config.botPassword,
    MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    {},
    credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        "OnTurnError Trace",
        `${error}`,
        "https://www.botframework.com/schemas/error",
        "TurnError"
    );

    // Send a message to the user
    await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Init Storage
const memoryStorage = new MemoryStorage();
// initialise the conversation state
export const conversationState = new ConversationState(memoryStorage);
// initialise the user state
const userState = new UserState(memoryStorage);

// Create the bot that will handle incoming messages.
const bot = new TeamsBot(conversationState, userState);

// Create the bot that will handle incoming messages.
//const searchApp = new SearchApp();

// Get initial settings from environment variables
const env = process.env.NODE_ENV || 'production';
info("NODE_ENV: " + env);

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
    info(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
// server.post("/api/messages", async (req, res) => {
//   await adapter.process(req, res, async (context) => {
//     await searchApp.run(context);
//   });
// });
server.post("/api/messages", async (req, res) => {
    await adapter.process(req, res, async (context) => {
        //await searchApp.run(context);
        await bot.run(context);
    }).catch((err) => {
        // Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, sholdn't throw this error.
        if (!err.message.includes("412")) {
            throw err;
        }
    })
});
server.get('/api/getSettings/:upn', async (req, res) => {
    info(`reading settings from ${path.join(__dirname, `/settings.${env}.json`)}`);
    const data = fs.readFileSync(path.join(__dirname, `/settings.${env}.json`));
    //const settings = require(`./settings.${env}.json`);
    const settings = JSON.parse(data)
    let ret = new Array<any>();
    settings.forEach(s => {
        if (s["owner"] == null || s["owner"].toLowerCase() == req.params.upn.toLowerCase()) {
            ret.push(_.omit(s, "owner"));
        }
    });
    res.send(200, ret);
});
server.post('/api/refreshtoken', async (req, res) => {
    const authService = new ServerSideAuthService(config.azureDevOpsProviderConfig);
    let token;
    let errorMessage = "";
    try {
        token = await authService.refreshToken(req.body.refreshToken);
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
        info(`gettokenbyauthcode authcode: ${req.body.code}; host: ${req.header('Host')}`);
        token = await authService.getTokenByCode(req.body.code);
        info(`gettokenbyauthcode-token: ${token}`);
        res.send(200, token);
    } catch (err) {
        err("/api/gettokenbyauthcode", err);
        res.send(400, {
            error: err
        });
    }
});
//server.post('/api/TaskReport', authenticateKey, async (req, res) => {
server.post('/api/TaskReport', authenticateBearKey, async (req: Request, res: Response) => {
    const dbService = new CompletedTasksDbSerivce();
    try {
        //info("body", req.body);
        const result = await dbService.addTask(req.params.upn, req.body.tasks, req.body.reportDate);
        //info("result", result);
        if (result && (result.statusCode == 200 || result.statusCode == 201))
            res.send(result.statusCode, { "id": result.item.id });
        else
            res.send(400, "Failed to save the task report.");
    } catch (err) {
        console.error(err);
        res.send(500, {
            error: err
        });
    }
});

server.get(
    //"/auth-:name(start|end|config).html",
    "/web/*",
    restify.plugins.serveStatic({
        //directory: path.join(__dirname, "public"),
        directory: path.join(__dirname),
    })
);
