import { Container, CosmosClient, CosmosClientOptions } from "@azure/cosmos";
import config from "../../config";

export class BaseDBService {
    public async getDbContainer(containerId: string): Promise<Container> {
        const options: CosmosClientOptions = {
            endpoint: config.cosmosDbConfig.CosmosDbEndPoint!,
            key: config.cosmosDbConfig.CosmosDbKey,
            userAgentSuffix: "M365PODDevOpsDB"
        };
        const client = new CosmosClient(options);
        const db = await client.database(config.cosmosDbConfig.CosmosDbId!);
        const container = await db.container(containerId);
        return container;
    }

}