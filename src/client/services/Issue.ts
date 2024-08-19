import moment from "moment";
import { IComboBoxOption, IDropdownOption, IGroup, SelectableOptionMenuItemType } from "@fluentui/react";
import { Area } from "./Area";
import { ISettingItem } from "./SettingService";
import * as _ from "lodash";

export class Issue {

    constructor(public id: number,
        public title: string,
        public areaId: number,
        public areaPath: string,
        public nodeName: string,
        public iterationId: number,
        public state: string,
        public url: string,
        public description: string,
        public axisCode: string,
        public createdDate: Date,
        public settingItemId: string) {

    }
    public get groupKey(): string {
        return `${this.settingItemId}-${this.areaPath}`;
    }
    /**
     * createIssueFromResponse
     */
    public static createIssueFromResponse(settingItemId: string, responseItem: any) {
        return new Issue(responseItem["id"],
            responseItem["fields"]["System.Title"],
            +(responseItem["fields"]["System.AreaId"]),
            responseItem["fields"]["System.AreaPath"],
            responseItem["fields"]["System.NodeName"],
            +(responseItem["fields"]["System.IterationId"]),
            responseItem["fields"]["System.State"],
            responseItem["url"],
            responseItem["fields"]["System.Description"],
            responseItem["fields"]["Custom.AxisCode"],
            responseItem["fields"]["System.CreatedDate"],
            settingItemId);
    }
    public static createIssueFromIssue(issue: Issue) {
        return new Issue(
            issue.id,
            issue.title,
            issue.areaId,
            issue.areaPath,
            issue.nodeName,
            issue.iterationId,
            issue.state,
            issue.url,
            issue.description,
            issue.axisCode,
            issue.createdDate,
            issue.settingItemId);
    }
    public toIDropdownOption(): IDropdownOption {
        return {
            key: this.id,
            text: this.title
        }
    }

}
export class ProjectIssues {
    private static CACHE_ISSUES = "AzureDevOps-AllIssues";
    private static _instance: ProjectIssues = null;

    constructor(public cachedIssues: IssueCollection[]) {

    }

    public static getSettingItemIssuesFromCache(settingItem: ISettingItem): IssueCollection {
        let ret: IssueCollection = null;
        const instance = ProjectIssues.getInstance();

        if (instance) {
            const existIssues = instance.cachedIssues.find(issues => issues.settingItemId == settingItem.id);
            ret = _.cloneDeep(existIssues);
        }

        return ret;
    }
    private static getInstance(): ProjectIssues {
        if (!this._instance) {
            let cacheJson = localStorage.getItem(ProjectIssues.CACHE_ISSUES);
            if (cacheJson) {
                this._instance = ProjectIssues.createIssuesFromJson(cacheJson);
            }
        }
        return this._instance;
    }
    private static createIssuesFromJson(json: string): ProjectIssues {
        let ret: ProjectIssues;
        let temp: ProjectIssues = JSON.parse(json) as ProjectIssues;
        if (temp) {
            let retIssues: IssueCollection[] = [];
            temp.cachedIssues.forEach(issues => {
                let tempIssues: Issue[] = [];
                issues.items.forEach(item => {
                    tempIssues.push(Issue.createIssueFromIssue(item));
                });
                retIssues.push(new IssueCollection(tempIssues, issues.settingItemId, issues.timeStamp));
            });
            ret = new ProjectIssues(retIssues);
        }
        return ret;
    }
    public static updateSettingItemIssuesToCache(issues: IssueCollection) {
        let instance: ProjectIssues = this.getInstance();
        if (instance) {
            const index = instance.cachedIssues.findIndex(item => item.settingItemId == issues.settingItemId);
            if (index == -1) {
                instance.cachedIssues.push(issues);
            } else {
                instance.cachedIssues[index] = issues;
            }
        } else {
            instance = new ProjectIssues([issues]);
        }
        localStorage.setItem(ProjectIssues.CACHE_ISSUES, JSON.stringify(instance));
    }
    public static updateSettingItemIssueToCache(issues: IssueCollection, issue: Issue) {
        let instance: ProjectIssues = this.getInstance();
        if (instance) {
            const issuesIndex = instance.cachedIssues.findIndex(item => item.settingItemId == issues.settingItemId);
            if (issuesIndex == -1) {
                instance.cachedIssues.push(issues);
            } else {
                //instance.cachedIssues[index] = issues;
                const issueIndex = instance.cachedIssues[issuesIndex].items.findIndex(item => item.id == issue.id);
                if (issueIndex == -1) {
                    instance.cachedIssues[issuesIndex].items.push(issue);
                } else {
                    instance.cachedIssues[issuesIndex].items[issueIndex] = issue;
                }
            }
        } else {
            instance = new ProjectIssues([issues]);
        }
        localStorage.setItem(ProjectIssues.CACHE_ISSUES, JSON.stringify(instance));
    }

}
export class IssueCollection {
    private static CACHE_ISSUES = "AzureDevOps-Issues";
    constructor(public items: Issue[], public settingItemId: string, public timeStamp: moment.Moment) {

    }
    /**
     * createIssuesFromResponse
     */
    public static createIssuesFromResponse(settingItem: ISettingItem, response: any): IssueCollection {
        let ret = null;
        if (response && response.value) {
            ret = new IssueCollection([], settingItem.id, moment());
            response.value.forEach(itemResponse => {
                ret.items.push(Issue.createIssueFromResponse(settingItem.id, itemResponse));
            })
        }
        return ret;
    }
    public static createIssuesFromJson(json: string): IssueCollection {
        let ret: IssueCollection;
        let temp: IssueCollection = JSON.parse(json) as IssueCollection;
        if (temp) {
            ret = new IssueCollection([], temp.settingItemId, temp.timeStamp);
            temp.items.forEach(item => ret.items.push(Issue.createIssueFromIssue(item)));
        }
        return ret;
    }
    public static getIsuesFromCache(): IssueCollection {
        let ret: IssueCollection = null;
        let cacheJson = localStorage.getItem(IssueCollection.CACHE_ISSUES);
        if (cacheJson) {
            ret = IssueCollection.createIssuesFromJson(cacheJson);
        }
        return ret;
    }

    public isUpdateToDate(): boolean {
        // Set update-to-date is true within 24 hours
        const diffHour = moment().diff(this.timeStamp, 'hour');
        return diffHour < 24;
    }
    public appendIssues(issues: IssueCollection) {
        issues.items.forEach(item => this.items.push(item));
    }
    // public saveToCache() {
    //     if (this.items) {
    //         localStorage.setItem(IssueCollection.CACHE_ISSUES, JSON.stringify(this));
    //     }
    // }
    // public updateIssue(issue: Issue) {
    //     if (this.items) {
    //         const itemIndex = this.items.findIndex(item => item.id == issue.id && item.settingItemId == issue.settingItemId);
    //         if (itemIndex == -1) {
    //             this.items.push(issue);
    //         } else {
    //             this.items[itemIndex] = issue;
    //         }
    //         this.saveToCache();
    //     }
    // }
    public getIssueById(id: number): Issue {
        return this.items.find(item => item.id == id);
    }
    public getIssueByUrl(settingItem: ISettingItem, url: string): Issue {
        const urlSection = url.split("/");
        const searchIssueId = +urlSection[urlSection.length - 1];
        // the task's relation issue Url is different than the issue URL itself. change logic to parse the issueId from URL
        //return this.items.find(item => item.url == url && item.settingItemId == settingItem.id);
        return this.items.find(item => item.id == searchIssueId && item.settingItemId == settingItem.id);
    }
    public getIssuesByArea(settingItem: ISettingItem, area: Area): IssueCollection {
        let issues = this.items.filter(item => (item.areaId == area.Id || item.areaId == area.ParentId) && item.settingItemId == settingItem.id);
        //let issues = this.items.filter(item => (item.areaId == area.Id) && item.settingItemId == settingItem.id);
        issues = issues.sort((a, b) => a.areaPath.toLowerCase().localeCompare(b.areaPath.toLowerCase()));
        return new IssueCollection(issues, settingItem.id, null);
    }
    public getIssueAxisCodeIDropdownOptionsByArea(areaId: number, includeHeader: boolean = true): IComboBoxOption[] {
        let ret: IComboBoxOption[] = [];
        const issues = this.items.filter(item => item.areaId == areaId);
        if (issues && issues.length > 0) {
            let axisCodeIssues = _.groupBy(issues, (item: Issue): string => {
                //return `${item.Area.Name}-${item.Issue.title}`;
                return item.axisCode;
            });
            ret = [];
            for (let code in axisCodeIssues) {
                //groups.push(new ReportItemGroup(name, areaIssues[name]));
                if (includeHeader)
                    ret.push({ key: `Header${axisCodeIssues[code][0].id}`, text: axisCodeIssues[code][0].title, itemType: SelectableOptionMenuItemType.Header });
                ret.push({ key: code, text: code });
            }
        }
        return ret;
    }
    public sortByTitle() {
        this.items = this.items.sort((a, b) => { return a.title.toLowerCase().localeCompare(b.title.toLowerCase()); });
        //this.items = orderBy(this.items, "title", ["asc"])
    }
    public sortByArea() {
        this.items = this.items.sort((a, b) => { return a.areaPath.toLowerCase().localeCompare(b.areaPath.toLowerCase()); });
        //this.items = orderBy(this.items, ['area'], ['asc']);
    }
    public toIDropdownOption(): IDropdownOption[] {
        let ret: IDropdownOption[] = [];
        if (this.items) {
            ret = this.items.map(t => t.toIDropdownOption());
        }
        return ret;
    }
    public createGroupsFromIssues(): IGroup[] {
        let ret: IGroup[] = [];
        const sortedItems = this.items.sort((a, b) => { return a.areaPath.toLowerCase().localeCompare(b.areaPath.toLowerCase()); });

        sortedItems.forEach((item: Issue, index: number, array: Issue[]) => {
            let existGroup = ret.find(g => g.key == item.groupKey);
            if (!existGroup) {
                existGroup = {
                    key: item.groupKey,
                    name: `${item.settingItemId}-${item.nodeName}`,
                    startIndex: index,
                    count: 1,
                    isCollapsed: true,
                    level: 0
                };
                ret.push(existGroup);
            } else {
                existGroup.count++;
            }
        });
        return ret;
    }

}
