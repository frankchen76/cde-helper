import { AzureDevOpsTokenProvider } from "./AzureDevOpsTokenProvider";
import { HttpClientService } from "./HttpClientService";
import config from "../config";
import { SearchTaskCollection } from "./SearchTask";
import moment from "moment";

export interface ISearchTaskOptions {
    taskName?: string;
    customerName?: string;
    status?: string;
    creationDate?: string;
}
export class SearchTaskService extends HttpClientService {
    constructor(provider: AzureDevOpsTokenProvider) {
        super(provider);
    }
    public async getMe(): Promise<any> {
        const url = "https://app.vssps.visualstudio.com/_apis/profile/profiles/me?api-version=7.1-preview.3";
        const response = await this.get(url);
        return response;

    }
    public async searchTasks(searchOption: ISearchTaskOptions): Promise<SearchTaskCollection> {
        let ret = null;
        let url = `${config.azureDevOpsProviderConfig.projectUrl}/_apis/wit/wiql?api-version=7.0&$top=50`;
        let wiql = `select [System.Id] from WorkItems where [System.WorkItemType] = 'Task'`;
        if (searchOption.taskName) {
            wiql += ` and [System.Title] contains '${searchOption.taskName}'`;
        }
        if (searchOption.customerName) {
            wiql += ` and [System.AreaPath] = 'CDE02\\\\${searchOption.customerName}'`;
        }
        if (searchOption.status) {
            wiql += ` and [System.State] = '${searchOption.status}'`;
        }
        if (searchOption.creationDate) {
            try {
                const dateRagne: { s?: string, e?: string } = JSON.parse(searchOption.creationDate) as { s?: string, e?: string };
                if (dateRagne) {
                    if (dateRagne.s) {
                        wiql += ` and [System.CreatedDate] >= '${moment(dateRagne.s, "MM/DD/YYYY").format("YYYY-MM-DD")}'`;
                    }
                    if (dateRagne.e) {
                        wiql += ` and [System.CreatedDate] <= '${moment(dateRagne.e, "MM/DD/YYYY").format("YYYY-MM-DD")}'`;
                    }
                }
            } catch (error) {
                console.log("processed creationDate failed: ", error);
            }
        }
        // if (this.query.start) {
        //     ret += ` and [System.CreatedDate] >= '${moment(this.query.start).format("YYYY-MM-DD")}'`;
        // }
        // if (this.query.end) {
        //     ret += ` and [System.CreatedDate] <= '${moment(this.query.end).format("YYYY-MM-DD")}'`;
        // }
        wiql += ` order by [System.CreatedDate] desc`;
        let idsBody = {
            query: wiql
        };
        console.log("wiql", wiql);

        // query all tasks IDs
        const idsResponse = await this.post(url, idsBody);

        // query all tasks details
        const ids = idsResponse["workItems"].map(item => item["id"]);
        const batchBody = {
            "ids": idsResponse["workItems"].map(item => item["id"]),
            "$expand": "all"
        };
        if (ids && ids.length > 0) {
            url = `${config.azureDevOpsProviderConfig.projectUrl}/_apis/wit/workitemsbatch?api-version=7.0`;
            const tasksResponse = await this.post(url, batchBody);
            ret = SearchTaskCollection.createTasksFromResponse(tasksResponse);
        }

        return ret;
    }

}
