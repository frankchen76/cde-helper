import { OutlookItemJSON } from "./OutlookItem";

export enum MessageActionType {
    PopupEmail = "PopupEmail",
    UpdateCategory = "UpdateCategory"
}
// export class DialogMessage {
//     constructor(public actionType: MessageActionType, public msgContent) {

//     }
//     public static createInstance(json: string) {
//         const obj = JSON.parse(json) as DialogMessage;
//         return new DialogMessage(obj.actionType, obj.msgContent);
//     }
//     public toJson(): string {
//         return JSON.stringify(this);
//     }
// }
export class DialogMessage {
    constructor(public actionType: MessageActionType) {

    }
    public static createInstance(json: string) {
        const obj = JSON.parse(json) as DialogMessage;
        //return new DialogMessage(obj.actionType, obj.msgContent);
        return obj;
    }
    public toJson(): string {
        return JSON.stringify(this);
    }
}
export class DialogMessagePopupEmail extends DialogMessage {
    constructor(public actionType: MessageActionType, public outlookObj: OutlookItemJSON) {
        super(actionType);
    }
    public static createInstance(json: string) {
        const obj = JSON.parse(json) as DialogMessagePopupEmail;
        return new DialogMessagePopupEmail(obj.actionType, OutlookItemJSON.createInstance(JSON.stringify(obj.outlookObj)));
    }
}
export class DialogMessageUpdateCategory extends DialogMessage {
    constructor(public actionType: MessageActionType, public outlookObj: OutlookItemJSON, public category: string) {
        super(actionType);
    }
    public static createInstance(json: string) {
        const obj = JSON.parse(json) as DialogMessageUpdateCategory;
        return new DialogMessageUpdateCategory(obj.actionType, OutlookItemJSON.createInstance(JSON.stringify(obj.outlookObj)), obj.category);
    }
}