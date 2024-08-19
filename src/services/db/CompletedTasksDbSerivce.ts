import { Container, CosmosClient, CosmosClientOptions } from "@azure/cosmos";
import { BaseDBService } from "./BaseDBService";
import config from "../../config";

export interface ITaskItem {
    id: string,
    UPN: string,
    reportDate: string,
    tasks: []
}
export class CompletedTasksDbSerivce extends BaseDBService {
    public async addTask(upn: string, tasks: [], reportDate: string): Promise<any> {
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_CompletedTasks!);
        const item = {
            id: `${upn}-${reportDate}`,
            reportDate: reportDate,
            UPN: upn,
            tasks: tasks
        };
        const response = await container.items.upsert(item);
        return response;
        // console.log("response", response);
        // if (response.statusCode === 200) {
        //     return {
        //         "id": response.item.id,
        //     }
        // } else
        //     return null;
        // console.log("response", response);
        // console.log("response.item", response.item);
        //return response.item;
    }

}