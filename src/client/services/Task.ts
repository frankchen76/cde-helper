import { find } from "lodash";
import { IGroup } from "@fluentui/react";
import { Issue } from "./Issue";
import { OutlookItem, OutlookItemJSON } from "./OutlookItem";
import { ISettingItem } from "./SettingService";
import * as _ from "lodash";

export class Task {
    public id: number;
    public title: string;
    public state: string;
    public url: string;
    public urlHtml: string;
    public description: string;
    public area: string;
    public areaId: number;
    public createdDate: Date;
    public dueDate: Date;
    public issueUrl: string;
    public outlookMessage: OutlookItemJSON;
    public issue: Issue;
    public completedWork: number;
    public tags: string;
    public settingItem: ISettingItem;

    public get areaName(): string {
        return this.area.indexOf("\\") != -1 ? this.area.split("\\")[this.area.split("\\").length - 1] : ""
    }
    public get groupKey(): string {
        return `${this.area}-${this.settingItem.id}`;
    }
    public get groupName(): string {
        return `${this.settingItem.name}-${this.areaName}`;
    }

    public static createTaskFromResponse(settingItem: ISettingItem, taskResponse: any): Task {
        let ret = new Task();
        ret.id = taskResponse["id"];
        ret.title = taskResponse["fields"]["System.Title"];
        ret.state = taskResponse["fields"]["System.State"];
        ret.url = taskResponse["url"];
        ret.description = taskResponse["fields"]["System.Description"];
        ret.area = taskResponse["fields"]["System.AreaPath"];
        ret.areaId = taskResponse["fields"]["System.AreaId"];
        ret.createdDate = taskResponse["fields"]["System.CreatedDate"];
        ret.dueDate = taskResponse["fields"]["Microsoft.VSTS.Scheduling.DueDate"] ? new Date(taskResponse["fields"]["Microsoft.VSTS.Scheduling.DueDate"]) : null;
        ret.outlookMessage = OutlookItemJSON.createInstance(taskResponse["fields"]["Custom.OutlookMessageId"]);
        ret.completedWork = +taskResponse["fields"]["Microsoft.VSTS.Scheduling.CompletedWork"];
        ret.tags = taskResponse["fields"]["System.Tags"];
        //ret.issue= issue
        if (taskResponse["relations"]) {
            const parentIssue = find(taskResponse["relations"], r => {
                return r["attributes"] ? r["attributes"]["name"] == "Parent" : false;
            });
            ret.issueUrl = parentIssue ? parentIssue["url"] : null;
        }
        if (taskResponse["_links"] && taskResponse["_links"]["html"] && taskResponse["_links"]["html"]["href"]) {
            ret.urlHtml = taskResponse["_links"]["html"]["href"];
        }
        ret.settingItem = settingItem;
        return ret;
    }

}

export class TaskCollection {
    constructor(public items: Task[]) {

    }
    public static createTasksFromResponse(settingItem: ISettingItem, response: any): TaskCollection {
        let ret = null;
        if (response && response.value) {
            ret = new TaskCollection([]);
            response.value.forEach(itemResponse => {
                ret.items.push(Task.createTaskFromResponse(settingItem, itemResponse));
            })
        }
        return ret;
    }
    public appendTasks(tasks: TaskCollection) {
        if (tasks && tasks.items && tasks.items.length > 0) {
            this.items = _.concat(this.items, tasks.items);
        }
    }
    public sort() {
        this.items = this.items.sort((a, b) => { return a.area.toLowerCase().localeCompare(b.area.toLowerCase()); });
        //this.items = orderBy(this.items, ['area'], ['asc']);
    }
    public createGroupsFromTasks(): IGroup[] {
        let ret: IGroup[] = [];
        this.items.forEach((item: Task, index: number, array: Task[]) => {
            let existGroup = ret.find(g => g.key == item.groupKey);
            if (!existGroup) {
                existGroup = {
                    key: item.groupKey,
                    //name: item.area.indexOf("\\") != -1 ? item.area.split("\\")[item.area.split("\\").length - 1] : "",
                    name: item.groupName,
                    startIndex: index,
                    count: 1,
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