import { BaseService } from "./BaseService";
import { HttpClientService } from "./HttpClientService";
import { IAddTaskBody } from "./IAddTaskBody";
import { ISettingItem } from "./SettingService";

export class WorkItemService extends BaseService {
    protected async getWorkItemsFromQueryId(settingItem: ISettingItem, queryId: string, includedFields: string): Promise<any> {
        let ret = null;
        let url = `${settingItem.baseUrl}/_apis/wit/wiql/${queryId}/?api-version=6.0`;

        const queryResponse = await this._httpClientService.get(url);
        const ids = queryResponse["workItems"].map(item => item["id"]).join(",");
        const qsFields = includedFields && includedFields != "" ? `&${includedFields}` : "";
        if (ids != '') {
            url = `${settingItem.baseUrl}/_apis/wit/workitems?ids=${ids}${qsFields}&api-version=6.0`;
            ret = await this._httpClientService.get(url);
        }
        return ret;
    }
    protected async getWorkItemsFromWiql(settingItem: ISettingItem, wiql: string, includedFields: string): Promise<any> {
        let ret = null;
        let url = `${settingItem.baseUrl}/_apis/wit/wiql?api-version=6.0`;
        const body = {
            query: wiql
        };

        const queryResponse = await this._httpClientService.post(url, body);
        const ids = queryResponse["workItems"].map(item => item["id"]).join(",");
        const qsFields = includedFields && includedFields != "" ? `&${includedFields}` : "";
        if (ids != '') {
            url = `${settingItem.baseUrl}/_apis/wit/workitems?ids=${ids}${qsFields}&api-version=6.0`;
            ret = await this._httpClientService.get(url);
        }
        return ret;
    }
    protected async updateWorkItemState(settingItem: ISettingItem, workItem: { id: number, state: string }): Promise<void> {
        let body: IAddTaskBody[] = [
            {
                op: "add",
                path: "/fields/System.State",
                value: workItem.state
            }
        ];
        const url = `${settingItem.baseUrl}/_apis/wit/workitems/${workItem.id}?api-version=6.0`;
        const response = await this._httpClientService.patch(url, body, true);
    }


}