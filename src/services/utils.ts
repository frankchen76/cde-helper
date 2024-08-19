import { CardFactory, Attachment, TeamsChannelAccount, TeamsInfo, TurnContext } from "botbuilder";
// import ACData = require("adaptivecards-templating");

export class Utils {
    // Bind AdaptiveCard with data
    // static renderAdaptiveCard(rawCardTemplate: any, dataObj?: any): Attachment {
    //     const cardTemplate = new ACData.Template(rawCardTemplate);
    //     const cardWithData = cardTemplate.expand({ $root: dataObj });
    //     const card = CardFactory.adaptiveCard(cardWithData);
    //     return card;
    // }
    static async getTeamAccount(context: TurnContext): Promise<TeamsChannelAccount> {
        const ret: TeamsChannelAccount = await TeamsInfo.getMember(
            context,
            context.activity.from.id
        );
        return ret;
    }
    static cleanupParam(value: string): string {
        if (!value) {
            return "";
        } else {
            let result = value.trim();
            //result = result.split(',')[0];          // Remove extra data
            result = result.replace("*", "");       // Remove wildcard characters from Copilot
            return result;
        }
    }

}
