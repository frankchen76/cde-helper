import { BaseDBService } from "./BaseDBService";
import config from "../../config";
import { info } from "console";
import { ItemDefinition, ItemResponse } from "@azure/cosmos";

export class SettingsDbSerivce extends BaseDBService {
    public async getSettings(upn: string): Promise<any> {
        let ret: any = null;
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_UserSettings!);
        const querySpec = {
            query: 'SELECT * FROM UserSettings c where c.upn=@upn',
            parameters: [
                {
                    name: '@upn',
                    value: upn
                }
            ]
        }

        const { resources: results } = await container
            .items.query(querySpec)
            .fetchAll();
        info(`Get settings for upn: ${upn}; results:`, results);
        if (results && results.length > 0) {
            ret = results[0];
        }
        return ret;

    }
    public async saveSettings(upn: string, setting: any): Promise<ItemResponse<ItemDefinition>> {
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_UserSettings!);
        setting.upn = upn;
        const response = await container.items.upsert(setting);
        return response;

    }

}