import {
    TeamsActivityHandler,
    CardFactory,
    TurnContext,
    ActivityTypes,
    SigninStateVerificationQuery,
    MessagingExtensionQuery,
    MessagingExtensionResponse,
    AdaptiveCardInvokeValue,
    AdaptiveCardInvokeResponse,
    UserState,
    ConversationState,
    StatePropertyAccessor,
    InvokeResponse,
} from "botbuilder";
import config from "../config";
import { AzureDevOpsTokenProvider, BotStateTokenStore, IUserToken, IAuthCode, IAzureDevOpsProviderConfig } from "../services/AzureDevOpsTokenProvider";
import { SearchTaskService, ISearchTaskOptions } from "../services/SearchTaskService";
import * as ACData from "adaptivecards-templating";
import axios from "axios";
import * as querystring from "querystring";
import helloWorldCard from "../adaptiveCards/helloWorldCard.json";
import editCard from "../adaptiveCards/editCard.json";
import { Utils } from "../services/utils";
import { info } from "../services/log";

const USER_TOKEN_PROPERTY = "userTokenProperty";

export class TeamsBot extends TeamsActivityHandler {
    private userTokenAccessor: StatePropertyAccessor<IUserToken>;

    constructor(private conversationState: ConversationState,
        private userState: UserState) {
        super();

        this.userTokenAccessor = userState.createProperty<IUserToken>(USER_TOKEN_PROPERTY);

        this.onMessage(async (context, next) => {
            console.log("Running with Message Activity.");
            const removedMentionText = TurnContext.removeRecipientMention(context.activity);
            const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
            await context.sendActivity(`Echo: ${txt}`);
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id) {
                    await context.sendActivity(
                        `Hi there! I'm a Teams bot that will echo what you said to me.`
                    );
                    break;
                }
            }
            await next();
        });
    }
    async run(context: TurnContext) {
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);
    }

    // sample
    // public async handleTeamsMessagingExtensionQuery(
    //     context: TurnContext,
    //     query: any
    // ): Promise<any> {
    //     console.log("conifg:", config);

    //     const searchQuery = query.parameters[0].value;
    //     const response = await axios.get(
    //         `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
    //             text: searchQuery,
    //             size: 8,
    //         })}`
    //     );

    //     const attachments = [];
    //     response.data.objects.forEach((obj) => {
    //         const template = new ACData.Template(helloWorldCard);
    //         const card = template.expand({
    //             $root: {
    //                 name: obj.package.name,
    //                 description: obj.package.description,
    //             },
    //         });
    //         const preview = CardFactory.heroCard(obj.package.name);
    //         const attachment = { ...CardFactory.adaptiveCard(card), preview };
    //         attachments.push(attachment);
    //     });

    //     return {
    //         composeExtension: {
    //             type: "result",
    //             attachmentLayout: "list",
    //             attachments: attachments,
    //         },
    //     };
    // }

    public async handleTeamsMessagingExtensionQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResponse> {
        // eslint-disable-next-line no-secrets/no-secrets
        /**
         * User Code Here.
         * If query without token, no need to implement handleMessageExtensionQueryWithSSO;
         * Otherwise, just follow the sample code below to modify the user code.
         */
        //console.log('return handleMessageExtensionQueryWithSSO', config);

        // if no token found in user property, return auth
        const tokenStore = new BotStateTokenStore(this.userTokenAccessor, context);
        const tokenProvider = await AzureDevOpsTokenProvider.createInstance(tokenStore, config.azureDevOpsProviderConfig);

        const valueObj = context.activity.value;
        // process login if activity.value has state "state": "CancelledByUser"
        if (valueObj.state) {
            console.log("state found in activity's value, process login...", valueObj.state);
            const authCode = JSON.parse(valueObj.state) as IAuthCode;
            //const authCode = valueObj.state as IAuthCode;
            // init access token for current user
            await tokenProvider.initUserWithAuthCode(authCode);
        }

        // get user's token
        const devOpsUser = await tokenProvider.getUserToken();
        console.log("🔑access token: ", devOpsUser.accessToken);

        // check if token is valid, if not, return auth card
        if (!devOpsUser.HasToken) {
            console.log("🔒No AccessToken/refreshtoken in user state, return auth for authentication");
            const response = await this.getSignInResponseForMessageExtension(config.azureDevOpsProviderConfig, devOpsUser, context);
            return response;
            // await context.sendActivity({
            //     value: { status: 200, body: response },
            //     type: ActivityTypes.InvokeResponse,
            // });
            // return;
        }

        // Process query after get authenicated
        let taskName, customerName, status, creationDate;

        // For now we have the ability to pass parameters comma separated for testing until the UI supports it.
        // So try to unpack the parameters but when issued from Copilot or the multi-param UI they will come
        // in the parameters array.
        console.log("query: ", query);
        if (query.parameters.length === 1 && query.parameters[0]?.name === "taskName") {
            [taskName, customerName, status, creationDate] = (query.parameters[0]?.value.split(','));
        } else {
            taskName = Utils.cleanupParam(query.parameters.find((element) => element.name === "taskName")?.value);
            customerName = Utils.cleanupParam(query.parameters.find((element) => element.name === "customerName")?.value);
            status = Utils.cleanupParam(query.parameters.find((element) => element.name === "taskStatus")?.value);
            creationDate = Utils.cleanupParam(query.parameters.find((element) => element.name === "taskCreationDate")?.value);
        }
        console.log(`🔎 taskName=${taskName}, customerName=${customerName}, status=${status}, taskCreationDate=${creationDate}`);

        const service = new SearchTaskService(tokenProvider);
        const searchOption: ISearchTaskOptions = {
            taskName,
            customerName,
            status,
            creationDate,
        }
        const meResult = await service.searchTasks(searchOption)
        console.log("🔢meResult count: ", meResult?.items?.length || 0);

        // compose adaptvie card and preview card
        let attachments = [];
        if (meResult && meResult.items && meResult.items.length > 0) {
            meResult.items?.forEach(task => {
                const template = new ACData.Template(editCard);
                const card = template.expand({
                    $root: task,
                });
                const preview = CardFactory.heroCard(task.title);
                const attachment = { ...CardFactory.adaptiveCard(card), preview };
                attachments.push(attachment);
            });
        }
        return {
            composeExtension: {
                type: "result",
                attachmentLayout: "list",
                attachments,
            },
        };

        // // search azure devops. 
        // const thumbnailCard = CardFactory.thumbnailCard(devOpsUser.upn);
        // // Message Extension return the user profile info to user.
        // return {
        //     composeExtension: {
        //         type: "result",
        //         attachmentLayout: "list",
        //         attachments: [thumbnailCard],
        //     },
        // };

    }
    private async getSignInResponseForMessageExtension(config: IAzureDevOpsProviderConfig, devOpsUser: IUserToken, context: TurnContext): Promise<any> {
        const teamMember = await Utils.getTeamAccount(context);
        const scopesArray = this.getScopesArray(config.scopes);
        info("teamMember", teamMember);

        //const signInLink = `${config.loginUrl}?scope=${encodeURI(scopesArray.join(" "))}&clientId=${config.clientId}&state=${devOpsUser.state}`;
        const signInUrl = `${config.loginUrl}?clientId=${config.clientId}&tenantId=${config.tenantId}&scope=${config.scopes}&login_hint=${teamMember.userPrincipalName}&stamp=${new Date().getTime()}`;

        console.log('signInLink', signInUrl);
        return {
            composeExtension: {
                type: "auth",
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: signInUrl,
                            title: "Azure DevOps Auth",
                        },
                    ],
                },
            },
        };
    }
    private getScopesArray(scopes: string | string[]): string[] {
        const scopesArray: string[] = typeof scopes === "string" ? scopes.split(" ") : scopes;
        return scopesArray.filter((x) => x !== null && x !== "");
    }

    // process adaptive card invoke event
    public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse<any>> {
        let runEvents = true;
        info('onInvokeActivity', context.activity);
        try {
            if (context.activity.name === 'adaptiveCard/action') {
                switch (context.activity.value.action.verb) {
                    case 'ok': {
                        info('onInvokeActivity', 'ok clicked');
                        break;
                        //return actionHandler.handleTeamsCardActionUpdateStock(context);
                    }
                    case 'restock': {
                        info('onInvokeActivity', 'restock clicked');
                        break;
                        // return actionHandler.handleTeamsCardActionRestock(context);
                    }
                    case 'cancel': {
                        info('onInvokeActivity', 'cancel clicked');
                        break;
                        // return actionHandler.handleTeamsCardActionCancelRestock(context);
                    }
                    default:
                        runEvents = false;
                        return super.onInvokeActivity(context);
                }
            } else {
                runEvents = false;
                return super.onInvokeActivity(context);
            }
        } catch (err) {
            if (err.message === 'NotImplemented') {
                return { status: 501 };
            } else if (err.message === 'BadRequest') {
                return { status: 400 };
            }
            throw err;
        } finally {
            if (runEvents) {
                this.defaultNextEvent(context)();
            }
        }

    }

    // The user has chosen to accept the settings by pressing the)
    // public async handleTeamsMessagingExtensionConfigurationQuerySettingUrl(_context: TurnContext, _query: MessagingExtensionQuery): Promise<MessagingExtensionResponse> {
    //     console.log('return handleTeamsMessagingExtensionConfigurationQuerySettingUrl');
    //     const linkUrl = config.azureDevOpsProviderConfig.loginUrl.replace('auth-start.html', 'auth-config.html');
    //     return {
    //         composeExtension: {
    //             type: "config",
    //             suggestedActions: {
    //                 actions: [
    //                     {
    //                         type: "openUrl",
    //                         value: linkUrl,
    //                         title: "Settings",
    //                     },
    //                 ],
    //             },
    //         },
    //     };
    // }

    // // Overloaded function. Receives invoke activities with the name 'composeExtension/setting
    // protected override async handleTeamsMessagingExtensionConfigurationSetting(_context: TurnContext, _settings: any): Promise<void> {
    //     console.log('return handleTeamsMessagingExtensionConfigurationSetting');
    //     // When the user submits the settings page, this event is fired.
    //     if (_settings.state != null) {
    //         //await this.userConfigurationProperty.set(_context, _settings.state);
    //     }
    // }

    // protected handleTeamsSigninTokenExchange(_context: TurnContext, _query: SigninStateVerificationQuery): Promise<void> {
    //     console.log('return handleTeamsSigninTokenExchange', _context);
    //     return super.handleTeamsSigninTokenExchange(_context, _query);
    // }
    // protected handleTeamsSigninVerifyState(_context: TurnContext, _query: SigninStateVerificationQuery): Promise<void> {
    //     console.log('return handleTeamsSigninVerifyState', _context);
    //     return super.handleTeamsSigninVerifyState(_context, _query);
    // }

    // private getSignInResponseForMessageExtensionWithSilentAuthConfig(
    //     clientId: string,
    //     initiateLoginEndpoint: string,
    //     scopes: string | string[]
    // ): any {
    //     const scopesArray = this.getScopesArray(scopes);
    //     const signInLink = `${initiateLoginEndpoint}?scope=${encodeURI(scopesArray.join(" "))}&clientId=${clientId}`;
    //     return {
    //         composeExtension: {
    //             type: "silentAuth",
    //             suggestedActions: {
    //                 actions: [
    //                     {
    //                         type: "openUrl",
    //                         value: signInLink,
    //                         title: "Message Extension OAuth",
    //                     },
    //                 ],
    //             },
    //         },
    //     };
    // }
}
