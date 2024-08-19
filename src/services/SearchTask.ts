import { concat, find } from "lodash";
import { NodeHtmlMarkdown } from "node-html-markdown";

export class SearchTask {
    public id: number;
    public title: string;
    public state: string;
    public url: string;
    public description: string;
    public area: string;
    public areaId: number;
    public createdDate: Date;
    public dueDate: Date;
    public issueUrl: string;
    public completedWork: number;
    public tags: string;

    public get areaName(): string {
        return this.area.indexOf("\\") != -1 ? this.area.split("\\")[this.area.split("\\").length - 1] : ""
    }

    public static createTaskFromResponse(taskResponse: any): SearchTask {
        let ret = new SearchTask();
        ret.id = taskResponse["id"];
        ret.title = taskResponse["fields"]["System.Title"];
        ret.state = taskResponse["fields"]["System.State"];
        ret.url = taskResponse["_links"]["html"]["href"];
        ret.area = taskResponse["fields"]["System.AreaPath"];
        ret.areaId = taskResponse["fields"]["System.AreaId"];
        ret.createdDate = taskResponse["fields"]["System.CreatedDate"];
        ret.dueDate = taskResponse["fields"]["Microsoft.VSTS.Scheduling.DueDate"] ? new Date(taskResponse["fields"]["Microsoft.VSTS.Scheduling.DueDate"]) : null;
        ret.completedWork = +taskResponse["fields"]["Microsoft.VSTS.Scheduling.CompletedWork"];
        ret.tags = taskResponse["fields"]["System.Tags"];
        //ret.issue= issue
        if (taskResponse["relations"]) {
            const parentIssue = find(taskResponse["relations"], r => {
                return r["attributes"] ? r["attributes"]["name"] == "Parent" : false;
            });
            ret.issueUrl = parentIssue ? parentIssue["url"] : null;
        }
        const nhm = new NodeHtmlMarkdown(
            /* options (optional) */ {},
            /* customTransformers (optional) */ undefined,
            /* customCodeBlockTranslators (optional) */ undefined
        );

        // convert HTML to Markdown
        ret.description = nhm.translate(/* html */ taskResponse["fields"]["System.Description"]);
        return ret;
    }

}

export class SearchTaskCollection {
    constructor(public items: SearchTask[]) {

    }
    public static createTasksFromResponse(response: any): SearchTaskCollection {
        let ret = null;
        if (response && response.value) {
            ret = new SearchTaskCollection([]);
            response.value.forEach(itemResponse => {
                ret.items.push(SearchTask.createTaskFromResponse(itemResponse));
            })
        }
        return ret;
    }
    public appendTasks(tasks: SearchTaskCollection) {
        if (tasks && tasks.items && tasks.items.length > 0) {
            this.items = concat(this.items, tasks.items);
        }
    }
    public sort() {
        this.items = this.items.sort((a, b) => { return a.area.toLowerCase().localeCompare(b.area.toLowerCase()); });
        //this.items = orderBy(this.items, ['area'], ['asc']);
    }
}