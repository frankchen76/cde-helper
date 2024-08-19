import { Issue, IssueCollection, ProjectIssues } from "./Issue";
import { HttpClientService } from "./HttpClientService";
import { ISettingItem, Setting } from "./SettingService";
import { WorkItemService } from "./WorkItemService";
import { IAddTaskBody } from "./IAddTaskBody";
import { Area } from "./Area";
import { ProfileService } from "./ProfileService";

export interface IIssueService {
    getIssuesByQuery(settingItem: ISettingItem, reload: boolean): Promise<IssueCollection>;
    getIssueByUrl(settingItem: ISettingItem, issueUrl: string): Promise<Issue>;
    getIssueById(settingItem: ISettingItem, issueId: number): Promise<Issue>;
    getIssuesByArea(settingItem: ISettingItem, area: Area): Promise<IssueCollection>
    updateIssue(assigneeUpn: string, settingItem: ISettingItem, issue: Issue, prevIssue?: Issue): Promise<Issue>;
    updateIssueState(settingItem: ISettingItem, issue: Issue): Promise<void>;
}
export class IssueService extends WorkItemService {
    private static CACHE_ISSUES = "AzureDevOps-AllIssues";
    private _issueIncludedFields = "fields=System.Title,System.Description,System.State,System.AreaId,System.AreaPath,System.NodeName,System.IterationId,System.CreatedDate,Custom.AxisCode";

    /**
     *  getIssuesByQueryId
     */
    public async getIssuesByQuery(settingItem: ISettingItem, reload: boolean = false): Promise<IssueCollection> {
        let ret: IssueCollection = ProjectIssues.getSettingItemIssuesFromCache(settingItem);
        if (reload || !ret || !ret.isUpdateToDate()) {
            const itemsResponse = await this.getWorkItemsFromQueryId(settingItem, settingItem.issuesQueryId, this._issueIncludedFields);
            ret = IssueCollection.createIssuesFromResponse(settingItem, itemsResponse);
            if (ret) {
                ProjectIssues.updateSettingItemIssuesToCache(ret);
            }
        }
        return ret;

        // let ret: IssueCollection = IssueCollection.getIsuesFromCache();
        // if (reload || !ret || !ret.isUpdateToDate) {
        //     const itemsResponse = await this.getWorkItemsFromQueryId(settingItem, settingItem.issuesQueryId, this._issueIncludedFields);
        //     ret = IssueCollection.createIssuesFromResponse(settingItem, itemsResponse);
        //     if (ret)
        //         ret.saveToCache();
        // }
        // return ret;
    }
    public async getIssueByUrl(settingItem: ISettingItem, issueUrl: string): Promise<Issue> {
        let cachedIssues = await this.getIssuesByQuery(settingItem);
        let existIssue = null;
        if (cachedIssues) {
            existIssue = cachedIssues.getIssueByUrl(settingItem, issueUrl);
        }
        if (!existIssue) {
            const url = `${issueUrl}?${this._issueIncludedFields}&api-version=6.0`;
            const issueResponse = await this._httpClientService.get(url);
            existIssue = Issue.createIssueFromResponse(settingItem.id, issueResponse);
            if (cachedIssues) {
                // cachedIssues.updateIssue(existIssue);
                ProjectIssues.updateSettingItemIssueToCache(cachedIssues, existIssue);
            }
        }
        return existIssue;
    }
    public async getIssueById(settingItem: ISettingItem, issueId: number): Promise<Issue> {
        const url = `${settingItem.baseUrl}/_apis/wit/workitems/${issueId}?${this._issueIncludedFields}&api-version=6.0`;
        //const url = `${issueUrl}?${this._issueIncludedFields}&api-version=6.0`;
        const issueResponse = await this._httpClientService.get(url);
        const updatedIssue = Issue.createIssueFromResponse(settingItem.id, issueResponse);

        let cachedIssues = await this.getIssuesByQuery(settingItem);

        if (cachedIssues) {
            //cachedIssues.updateIssue(updatedIssue);
            ProjectIssues.updateSettingItemIssueToCache(cachedIssues, updatedIssue);
        }

        return updatedIssue;
    }
    public async getIssuesByArea(settingItem: ISettingItem, area: Area): Promise<IssueCollection> {
        const issues = await this.getIssuesByQuery(settingItem);
        return issues.getIssuesByArea(settingItem, area);
    }
    public async updateIssue(assigneeUpn: string, settingItem: ISettingItem, issue: Issue, prevIssue?: Issue): Promise<Issue> {
        let ret: Issue = null;
        let response = null;
        let url = "";
        //const assigneeUpn = await ProfileService.getUpn();

        if (issue.id == 0) {
            let body: IAddTaskBody[] = [
                {
                    op: "add",
                    path: "/fields/System.Title",
                    value: issue.title
                },
                {
                    op: "add",
                    path: "/fields/System.AreaId",
                    value: issue.areaId
                },
                {
                    op: "add",
                    path: "/fields/System.IterationId",
                    value: issue.iterationId
                },
                {
                    op: "add",
                    path: "/fields/System.AssignedTo",
                    value: assigneeUpn
                },
                {
                    op: "add",
                    path: "/fields/System.Description",
                    value: issue.description
                },
                {
                    "op": "add",
                    "path": "/fields/Custom.AxisCode",
                    "value": issue.axisCode
                }
            ];
            url = `${settingItem.baseUrl}/_apis/wit/workitems/$issue?api-version=6.0`;
            response = await this._httpClientService.post(url, body, true);

            // Update the state if state isn't "To Do" 
            if (issue.state != "To Do") {
                const newIssueId = response["id"];
                body = [];
                body.push({
                    op: "add",
                    path: "/fields/System.State",
                    value: issue.state
                });
                url = `${settingItem.baseUrl}/_apis/wit/workitems/${newIssueId}?api-version=6.0`;
                response = await this._httpClientService.patch(url, body, true);
            }

        } else {
            let body: IAddTaskBody[] = [
                {
                    op: "add",
                    path: "/fields/System.Title",
                    value: issue.title
                },
                {
                    op: "add",
                    path: "/fields/System.AreaId",
                    value: issue.areaId
                },
                {
                    op: "add",
                    path: "/fields/System.State",
                    value: issue.state
                },
                {
                    op: "add",
                    path: "/fields/System.Description",
                    value: issue.description
                },
                {
                    "op": "add",
                    "path": "/fields/Custom.AxisCode",
                    "value": issue.axisCode
                }
            ];

            url = `${settingItem.baseUrl}/_apis/wit/workitems/${issue.id}?api-version=6.0`;
            response = await this._httpClientService.patch(url, body, true);
        }
        const issueId = response["id"];
        ret = await this.getIssueById(settingItem, +issueId);
        return ret;
    }
    public async updateIssueState(settingItem: ISettingItem, issue: Issue): Promise<void> {
        await super.updateWorkItemState(settingItem, issue);
    }

}