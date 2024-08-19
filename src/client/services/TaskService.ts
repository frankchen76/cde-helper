import moment from "moment";
import { IAddTaskBody } from "./IAddTaskBody";
import { IIssueService, IssueService } from "./IssueService";
import { ProfileService } from "./ProfileService";
import { ISettingItem, Setting } from "./SettingService";
import { Task, TaskCollection } from "./Task";
import { WorkItemService } from "./WorkItemService";

export interface ITaskQuery {
    areaPath: string;
    start?: Date;
    end?: Date;
    status?: string;
}
export class TaskQuery {
    constructor(private query: ITaskQuery) {

    }
    public get wiql(): string {
        let ret = `select [System.Id] from WorkItems where [System.WorkItemType] = 'Task' and [System.AreaPath] = '${this.query.areaPath}'`;
        if (this.query.status) {
            ret += ` and [System.State] = '${this.query.status}'`;
        }
        if (this.query.start) {
            ret += ` and [System.CreatedDate] >= '${moment(this.query.start).format("YYYY-MM-DD")}'`;
        }
        if (this.query.end) {
            ret += ` and [System.CreatedDate] <= '${moment(this.query.end).format("YYYY-MM-DD")}'`;
        }
        return ret;
    }

}
export interface ITaskService {
    getTasksByQueryId(settingItem: ISettingItem): Promise<TaskCollection>;
    getTasksByArea(settingItem: ISettingItem, query: ITaskQuery): Promise<TaskCollection>;
    getTaskById(settingItem: ISettingItem, taskId: number): Promise<Task>;
    updateTask(assigneeUpn: string, settingItem: ISettingItem, task: Task, comments?: string, prevTask?: Task): Promise<Task>;
    updateTaskState(settingItem: ISettingItem, task: Task): Promise<void>;
}
export class TaskService extends WorkItemService implements ITaskService {
    private _taskIncludedFields = "fields=System.Title,System.State,System.Description,System.AreaPath,System.AreaId,System.CreatedDate,Custom.OutlookMessageId,Microsoft.VSTS.Scheduling.CompletedWork,System.Tags,Microsoft.VSTS.Scheduling.DueDate";
    public _issueService: IIssueService;
    public async getTasksByQueryId(settingItem: ISettingItem): Promise<TaskCollection> {
        //const { issueService } = this._serivceContext;

        let ret: TaskCollection = null;

        const includedFields = "$expand=all";
        //const itemsResponse = await this.getWorkItemsFromQueryId(settingItem, settingItem.tasksQueryId, this._taskIncludedFields);
        const itemsResponse = await this.getWorkItemsFromQueryId(settingItem, settingItem.tasksQueryId, includedFields);
        ret = TaskCollection.createTasksFromResponse(settingItem, itemsResponse);

        if (ret) {
            for (let item of ret.items) {
                if (item.issueUrl) {
                    item.issue = await this._issueService.getIssueByUrl(settingItem, item.issueUrl);
                }
            }
        }

        return ret;
    }
    // Get tasks by area path and date range using wiql
    public async getTasksByArea(settingItem: ISettingItem, query: ITaskQuery): Promise<TaskCollection> {
        let ret: TaskCollection = null;

        const includedFields = "$expand=all";
        const wiql = new TaskQuery(query).wiql;
        //const itemsResponse = await this.getWorkItemsFromQueryId(settingItem, settingItem.tasksQueryId, this._taskIncludedFields);
        const itemsResponse = await this.getWorkItemsFromWiql(settingItem, wiql, includedFields);
        ret = TaskCollection.createTasksFromResponse(settingItem, itemsResponse);

        if (ret) {
            for (let item of ret.items) {
                if (item.issueUrl) {
                    item.issue = await this._issueService.getIssueByUrl(settingItem, item.issueUrl);
                }
            }
        }

        return ret;
    }
    public async getTaskById(settingItem: ISettingItem, taskId: number): Promise<Task> {
        //const { issueService } = this._serivceContext;

        const url = `${settingItem.baseUrl}/_apis/wit/workitems/${taskId}?$expand=all&api-version=6.0`;
        const taskResponse = await this._httpClientService.get(url);
        let ret = await Task.createTaskFromResponse(settingItem, taskResponse);
        if (ret && ret.issueUrl) {
            ret.issue = await this._issueService.getIssueByUrl(settingItem, ret.issueUrl);
        }
        return ret;
    }

    public async updateTask(assigneeUpn: string, settingItem: ISettingItem, task: Task, comments?: string, prevTask?: Task): Promise<Task> {
        let ret: Task = null;
        let response = null;
        let url = "";
        //const assigneeUpn = await ProfileService.getUpn();

        if (task.id == 0) {
            let body: IAddTaskBody[] = [
                {
                    op: "add",
                    path: "/fields/System.Title",
                    value: task.title
                },
                {
                    op: "add",
                    path: "/fields/System.AreaId",
                    value: task.areaId
                },
                {
                    op: "add",
                    path: "/fields/System.IterationId",
                    value: task.issue.iterationId
                },
                {
                    op: "add",
                    path: "/relations/-",
                    value: {
                        rel: "System.LinkTypes.Hierarchy-Reverse",
                        url: task.issue.url,
                        attributes: {
                            isLocked: false,
                            name: "Parent"
                        }
                    }
                },
                {
                    op: "add",
                    path: "/fields/Custom.OutlookMessageId",
                    value: task.outlookMessage.toJson()
                },
                {
                    op: "add",
                    path: "/fields/Microsoft.VSTS.Scheduling.CompletedWork",
                    value: task.completedWork
                },
                {
                    op: "add",
                    path: "/fields/System.AssignedTo",
                    value: assigneeUpn
                },
                {
                    op: "add",
                    path: "/fields/System.Description",
                    value: task.description
                },
                {
                    "op": "add",
                    "path": "/fields/System.Tags",
                    "value": task.tags
                }
            ];
            if (task.dueDate) {
                body.push({
                    "op": "add",
                    "path": "/fields/Microsoft.VSTS.Scheduling.DueDate",
                    "value": task.dueDate
                });
            }
            url = `${settingItem.baseUrl}/_apis/wit/workitems/$task?api-version=6.0`;
            response = await this._httpClientService.post(url, body, true);

            // Update the state if state isn't "To Do" 
            if (task.state != "To Do") {
                const newTaskId = response["id"];
                body = [];
                body.push({
                    op: "add",
                    path: "/fields/System.State",
                    value: task.state
                });
                url = `${settingItem.baseUrl}/_apis/wit/workitems/${newTaskId}?api-version=6.0`;
                response = await this._httpClientService.patch(url, body, true);
            }
        } else {
            let body: IAddTaskBody[] = [];
            if (task.title != prevTask.title) {
                body.push({
                    op: "add",
                    path: "/fields/System.Title",
                    value: task.title
                });
            }
            if (task.areaId != prevTask.areaId) {
                body.push({
                    op: "add",
                    path: "/fields/System.AreaId",
                    value: task.areaId
                });
            }
            if (task.issue.iterationId != prevTask.issue.iterationId) {
                body.push({
                    op: "add",
                    path: "/fields/System.IterationId",
                    value: task.issue.iterationId
                });
            }
            if (task.completedWork != prevTask.completedWork) {
                body.push({
                    op: "add",
                    path: "/fields/Microsoft.VSTS.Scheduling.CompletedWork",
                    value: task.completedWork
                });
            }
            if (task.state != prevTask.state) {
                body.push({
                    op: "add",
                    path: "/fields/System.State",
                    value: task.state
                });
            }
            if (task.description != prevTask.description) {
                body.push({
                    op: "add",
                    path: "/fields/System.Description",
                    value: task.description
                });
            }
            body.push({
                "op": "add",
                "path": "/fields/System.Tags",
                "value": task.tags
            });

            // let body: IAddTaskBody[] = [
            //     {
            //         op: "add",
            //         path: "/fields/System.Title",
            //         value: task.title
            //     },
            //     {
            //         op: "add",
            //         path: "/fields/System.AreaId",
            //         value: task.areaId
            //     },
            //     {
            //         op: "add",
            //         path: "/fields/System.IterationId",
            //         value: task.issue.iterationId
            //     },
            //     {
            //         op: "add",
            //         path: "/fields/Microsoft.VSTS.Scheduling.CompletedWork",
            //         value: task.completedWork
            //     },
            //     {
            //         op: "add",
            //         path: "/fields/System.State",
            //         value: task.state
            //     },
            //     {
            //         op: "add",
            //         path: "/fields/System.Description",
            //         value: task.description
            //     },
            //     {
            //         "op": "add",
            //         "path": "/fields/System.Tags",
            //         "value": task.tags
            //     }
            // ];
            if (task.dueDate && task.dueDate != prevTask.dueDate) {
                body.push({
                    "op": "add",
                    "path": "/fields/Microsoft.VSTS.Scheduling.DueDate",
                    "value": task.dueDate
                });
            }

            url = `${settingItem.baseUrl}/_apis/wit/workitems/${task.id}?api-version=6.0`;
            response = await this._httpClientService.patch(url, body, true);

            //update issue if different. 1. remove the existing parent. 2: add new one
            if (task.issue.id != prevTask.issue.id) {
                body = [
                    {
                        op: "remove",
                        path: "/relations/0"
                    },
                    {
                        op: "add",
                        path: "/relations/-",
                        value: {
                            rel: "System.LinkTypes.Hierarchy-Reverse",
                            url: task.issue.url,
                            attributes: {
                                isLocked: false,
                                name: "Parent"
                            }
                        }
                    }
                ];
                response = await this._httpClientService.patch(url, body, true);
            }
        }

        // update comments
        if (comments) {
            url = `${settingItem.baseUrl}/_apis/wit/workitems/${task.id}/comments?api-version=7.2-preview.4`;
            const commentBody = {
                "text": comments
            };
            const commentResponse = this._httpClientService.post(url, commentBody);
        }

        const taskId = response["id"];
        ret = await this.getTaskById(settingItem, +taskId);
        return ret;
    }
    public async updateTaskState(settingItem: ISettingItem, task: Task): Promise<void> {
        await super.updateWorkItemState(settingItem, task);
    }
}
