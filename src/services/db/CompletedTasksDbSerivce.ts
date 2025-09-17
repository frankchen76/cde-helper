import { Container, CosmosClient, CosmosClientOptions } from "@azure/cosmos";
import { BaseDBService } from "./BaseDBService";
import config from "../../config";
import { info } from "../log";
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
    }
    public async getTasks(upn: string, reportDate: string): Promise<any> {
        const container = await super.getDbContainer(config.cosmosDbConfig.CosmosDbContainerId_CompletedTasks!);
        const querySpec = {
            //query: `SELECT * FROM c where c.UPN='tachen@microsoft.com' and c.reportDate='2025-09-11'`
            query: `SELECT * FROM c where c.UPN=@u and c.reportDate=@rd`,
            parameters: [
                { name: "@u", value: upn.trim() },
                { name: "@rd", value: reportDate.trim() }
            ]
        }
        //info("getTasks-result", querySpec.query);
        const result = await container.items.query(querySpec).fetchAll();
        //info("getTasks-result", result);
        return result;
    }

}