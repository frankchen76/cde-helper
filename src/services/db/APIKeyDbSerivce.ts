import { CosmosClient, CosmosClientOptions } from "@azure/cosmos";
import { BaseDBService } from "./BaseDBService";
import config from "../../config";
import { info } from "../log";

export interface IAPIKeyItem {
    id: string,
    upn: string
}
export class APIKeyDbService extends BaseDBService {
    public async authenticateEmail(email: string): Promise<boolean> {
        let ret = false;
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_APIKeys!);

        const querySpec = {
            query: 'SELECT a.UPN, a.id FROM APIKeys a WHERE a.UPN=@email',
            parameters: [
                {
                    name: '@email',
                    value: email
                }
            ]
        }

        const { resources: results } = await container
            .items.query(querySpec)
            .fetchAll();
        ret = results && results.length > 0;
        return ret;
    }
    public async authenticateKey(key: string): Promise<IAPIKeyItem | null> {
        let ret: IAPIKeyItem | null = null;
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_APIKeys!);

        const querySpec = {
            query: 'SELECT a.UPN, a.id FROM APIKeys a WHERE a.apiKey=@apikey',
            parameters: [
                {
                    name: '@apikey',
                    value: key
                }
            ]
        }

        const { resources: results } = await container
            .items.query(querySpec)
            .fetchAll();

        if (results && results.length > 0) {
            ret = {
                id: results[0].id,
                upn: results[0].UPN
            };
        }
        return ret;
    }

}