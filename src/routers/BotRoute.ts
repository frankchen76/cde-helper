import * as restify from "restify";
import {
    CloudAdapter,
    ConfigurationServiceClientCredentialFactory,
    ConfigurationBotFrameworkAuthentication,
    TurnContext,
    MemoryStorage,
    ConversationState,
    UserState,
} from "botbuilder";
import config from "../config";
import { err, info } from '../services/log';
import { TeamsBot } from "../bot/teamsBot";
const path = require("path");

export const BotRoute = (server) => {
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
    const conversationState = new ConversationState(memoryStorage);
    // initialise the user state
    const userState = new UserState(memoryStorage);

    // Create the bot that will handle incoming messages.
    const bot = new TeamsBot(conversationState, userState);

    // Get initial settings from environment variables
    const env = process.env.NODE_ENV || 'production';
    info("NODE_ENV: " + env);
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

}