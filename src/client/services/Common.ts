import { IDropdownOption, IStackTokens } from "@fluentui/react";
import moment from "moment";

export enum TaskFormModeEnum {
    CreateTask,
    CreateEmailTask,
    UpdateTask
};
export enum IssueFormModeEnum {
    CreateIssue,
    UpdateIssue
};
export enum StateEnum {
    ToDo = "To Do",
    Doing = "Doing",
    Done = "Done"
};
export interface RestAPIError {
    $id: string;
    innerException: any;
    message: string;
    typeName: string;
    typeKey: string;
    errorCode: number;
    eventId: number;
}
export class Common {
    public static CONTAINER_STACK_TOKENS: IStackTokens = { childrenGap: 5 };
    public static CATEGORIES = [StateEnum.ToDo.toString(), StateEnum.Doing.toString(), StateEnum.Done.toString()];

    public static getTaskFormModeFromId(id: number) {
        let ret: TaskFormModeEnum = TaskFormModeEnum.CreateEmailTask;
        switch (id) {
            case -1:
                ret = TaskFormModeEnum.CreateTask;
                break;
            case 0:
                ret = TaskFormModeEnum.CreateEmailTask;
                break;
            default:
                ret = TaskFormModeEnum.UpdateTask;
                break;
        }
        return ret;
    }
    public static getIssueFormModeFromId(id: number): IssueFormModeEnum {
        return id == 0 ? IssueFormModeEnum.CreateIssue : IssueFormModeEnum.UpdateIssue;
    }
    public static getStateOptions(): IDropdownOption[] {
        return [
            { key: StateEnum.ToDo, text: StateEnum.ToDo },
            { key: StateEnum.Doing, text: StateEnum.Doing },
            { key: StateEnum.Done, text: StateEnum.Done },
        ];
    }
    public static getErrorMessage(error: any, fromRestAPI: boolean = true): string {
        let ret = error.toString();
        if (fromRestAPI) {
            try {
                const apiError = JSON.parse(error.toString()) as RestAPIError;
                if (apiError) {
                    ret = apiError.message;
                }
            } catch (error) {

            }
        }
        return ret;
    }
    public static dateToDurationString(date: Date, id?: number): string {
        let ret = "Unknown";
        if (date) {
            const localDate = moment(date).local();
            const created = moment(localDate.format(), "YYYY-MM-DD");
            var today = moment().startOf("day");
            //console.log(`${id | 0}: today: ${today.format()}, created: ${created.format()}, date: ${localDate}`);
            const days = today.diff(created, "days");
            const diff = Math.abs(days);
            if (diff == 0) {
                ret = "Today";
            } else if (diff == 1) {
                ret = days > 0 ? "Yesterday" : "Tomorrow";
            } else if (diff <= 7) {
                ret = `${diff} days ${days > 0 ? "ago" : "later"}`
            } else {
                ret = created.format("MM/DD/yyyy")
            }
        }
        return ret;
    }
}
export class ExecutingResult {
    constructor(public isRunning: boolean,
        public displayMessage: boolean,
        public message: string = "",
        public isError: boolean = false) {

    }
    public static createInstance(): ExecutingResult {
        return new ExecutingResult(false, false);
    }
    public static start(): ExecutingResult {
        return new ExecutingResult(true, false);
    }
    public static complete(displayMessage: boolean = false, message: string = "", isError: boolean = false): ExecutingResult {
        return new ExecutingResult(false, displayMessage, message, isError);
    }
}

