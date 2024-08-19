import { ISettingItem, Setting } from "./SettingService";
import { HistoryItemCollection, ReportItem, ReportItemCollection } from "./ReportItem";
import { IssueService } from "./IssueService";
import { WorkItemService } from "./WorkItemService";
import { find } from "lodash";
import { ApiKeyAuthHeader, HttpClientService } from "./HttpClientService";
import moment from "moment";

export interface IReportService {
    //getReportItems(settingItem: ISettingItem): Promise<ReportItemCollection>;
    getReportItems(setting: Setting): Promise<ReportItemCollection>;
}
export class ReportService extends WorkItemService implements IReportService {
    private _reportTaskIncludeFields = "$expand=All";

    public async getReportItems(setting: Setting): Promise<ReportItemCollection> {
        let ret: ReportItemCollection = null;
        for (const settingItem of setting.items) {
            const reportItems = await this.getReportItemsForSettingItem(settingItem);

            if (reportItems) {
                if (ret == null) {
                    ret = reportItems;
                } else {
                    ret.addReportItems(reportItems.Items);
                }
                //console.log(`settingItem: ${settingItem.id}; Report items count: ${ret.Items.length}`);
            }

        }
        // Log report items to DB
        if (ret && ret.Items && ret.Items.length > 0) {
            await this.logReportItemsToDb(setting.apiKey, ret);
        }
        return ret;
    }
    private async getReportItemsForSettingItem(settingItem: ISettingItem): Promise<ReportItemCollection> {
        let ret: ReportItemCollection = new ReportItemCollection();
        const issueService = new IssueService();

        const itemsResponse = await this.getWorkItemsFromQueryId(settingItem, settingItem.reportQueryId, this._reportTaskIncludeFields);
        if (itemsResponse && itemsResponse.value) {
            for (let jsonItem of itemsResponse.value) {
                let issue = null;
                if (jsonItem["relations"]) {
                    const parentIssue = find(jsonItem["relations"], r => {
                        return r["attributes"] ? r["attributes"]["name"] == "Parent" : false;
                    });
                    const url: string = parentIssue ? parentIssue["url"] : null;
                    issue = await issueService.getIssueByUrl(settingItem, url);
                }

                let reportItem = ReportItem.createInstanceFromJSON(jsonItem);
                reportItem.Issue = issue;

                if (reportItem.IsCreatedSameAsClosed()) {
                    //if the task is created and closed at the same day. 
                    reportItem.TodayHours = reportItem.CompletedWork;
                } else {
                    //get complete hour based on history. 
                    await this.loadReportItemHistory(settingItem, reportItem);
                }
                if (reportItem.TodayHours > 0)
                    ret.Items.push(reportItem);
            }
            ret.Items.sort();
        }

        return ret;
    }
    private async logReportItemsToDb(apiKey: string, reportItems: ReportItemCollection): Promise<string> {
        //const httpClientServiceWithApiKey = new HttpClientService(new ApiKeyAuthHeader(apiKey));
        const url = `${location.protocol}//${location.host}/api/TaskReport`;
        const body = {
            "reportDate": moment().format("YYYY-MM-DD"),
            "tasks": reportItems.Items
        };
        //const result = await httpClientServiceWithApiKey.post(url, body);
        const result = await this._httpClientService.post(url, body);
        return result.id;
    }
    private async loadReportItemHistory(settingItem: ISettingItem, reportItem: ReportItem): Promise<void> {

        let url = `${settingItem.baseUrl}/_apis/wit/workitems/${reportItem.Id}/revisions?$expand=All&api-version=6.0`;
        const historyResponse = await this._httpClientService.get(url);
        const histories = HistoryItemCollection.createInstanceFromJSON(historyResponse);
        reportItem.TodayHours = histories.getSelectedDateHour();
    }

}