import { findIndex, forEach, indexOf, isArray } from 'lodash';
import moment from "moment";
import { Common } from './Common';
import { DialogMessage, DialogMessagePopupEmail, DialogMessageUpdateCategory, MessageActionType } from './DialogMessage';
import { error, info } from './log';
import { jwtDecode } from "jwt-decode";


class OwsTokenHelper {
    private static owsCallbackToken = null;

    private static async getCallbackTokenAsync(): Promise<string> {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
                if (result.status == Office.AsyncResultStatus.Succeeded)
                    resolve(result.value);
                else {
                    error(`Error in getCallbackTokenAsync`, result.error);
                    reject(result.error);
                }
            });
        });
    }

    public static async getOwsToken(): Promise<string> {
        let ret = null;
        try {
            if (!OwsTokenHelper.owsCallbackToken) {
                info(`Getting office callback token because of empty.`);
                OwsTokenHelper.owsCallbackToken = await OwsTokenHelper.getCallbackTokenAsync();

                //start a timer to refresh the token every 60 seconds
                setInterval(async () => {
                    try {
                        const decodedToken = jwtDecode(OwsTokenHelper.owsCallbackToken);
                        const tokenExpired = moment(new Date(decodedToken.exp * 1000));
                        const now = moment();
                        info(`Checking if Office callback token is expired. tokenExpired: ${tokenExpired}, now: ${now}, token: ${OwsTokenHelper.owsCallbackToken}`);
                        //const tokenValid = tokenExpired > moment.utc();
                        if (tokenExpired.isBefore(now)) {
                            info(`Office callback token was expired, refresh it.`);
                            OwsTokenHelper.owsCallbackToken = await OwsTokenHelper.getCallbackTokenAsync();
                        } else {
                            info(`Office callback token is still valid.`);
                        }
                    } catch (err) {
                        error(`Error in token refresh timer-getOwsToken:`, err);
                    }
                }, 2700000); // refresh token by every 45 minutes
            }

            const decodedToken = jwtDecode(OwsTokenHelper.owsCallbackToken);
            const tokenExpired = moment(new Date(decodedToken.exp * 1000));
            const tokenValid = tokenExpired > moment();
            info(`tokenExpired: ${tokenExpired}, tokenValid: ${tokenValid}`);
            if (!tokenValid) {
                info(`Getting office callback token because of expiration.`);
                OwsTokenHelper.owsCallbackToken = await OwsTokenHelper.getCallbackTokenAsync();
            }
            ret = OwsTokenHelper.owsCallbackToken;
        } catch (err) {
            error("error in getOwsToken", err);
        }
        return ret;

    }
}

export class OutlookItem {
    private _item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead;
    private _itemId: string;
    private _itemType: string;
    private _subject: string;
    private _start: moment.Moment;
    private _end: moment.Moment;
    private _body: string;
    private _froms: string;
    private _categories: string[];

    public get ItemId(): string { return this._itemId; }
    public get ItemType(): string { return this._itemType; }
    public get Subject(): string { return this._subject; }
    public get Body(): string { return this._body; }
    public get Start(): moment.Moment { return this._start; }
    public get End(): moment.Moment { return this._end; }
    public get Duration(): moment.Duration {
        let ret: moment.Duration = null;
        if (this._start && this._end)
            ret = moment.duration(this._end.diff(this._start));
        return ret;
    }

    public IsSame(item: OutlookItem): boolean {
        return this.ItemId == item.ItemId;
    }
    public toOutlookItemJSON() {
        return new OutlookItemJSON(this.ItemId, this.ItemType);
    }
    public emailDomainExists(emailDomains: string[]): boolean {
        let ret = false;
        if (this._froms) {
            for (const emailDomain of emailDomains) {
                ret = this._froms.toLowerCase().indexOf(emailDomain.toLowerCase()) != -1;
                if (ret) break;
            }
        }
        return ret;
    }
    public areaExistInSubject(areaName: string): boolean {
        return this._subject.toLowerCase().indexOf(areaName.toLowerCase()) != -1;
    }
    public areaExistInCategories(areaName: string): boolean {
        let ret = false;
        if (this._categories) {
            ret = this._categories.findIndex(c => c.toLowerCase().indexOf(areaName.toLowerCase()) != -1) != -1;
        }
        return ret;
    }

    public static async createInstance(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<OutlookItem> {
        let ret: OutlookItem = null;
        try {
            ret = new OutlookItem();
            // console.log("outlookitem");
            // console.log(item);
            ret._item = item;
            ret._itemType = item.itemType;
            ret._itemId = await OutlookItem._getItemId(item);
            ret._subject = await OutlookItem._getItemSubject(item);
            ret._body = await OutlookItem._getItemBody(item);
            if (item.itemType == "appointment") {
                ret._start = await OutlookItem._getItemStart(item);
                ret._end = await OutlookItem._getItemEnd(item);
            }
            if (item.itemType == "message") {
                ret._froms = await OutlookItem._getItemFrom(item);
            }
            ret._categories = await OutlookItem._getItemCategories(item);

        } catch (error) {
            console.error(error);
        }
        return ret;
    }
    private static async _getItemId(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<string> {
        return new Promise((resolve, reject) => {
            if (item.internetMessageId) {
                resolve(item.internetMessageId);
            }
            else if (item.itemId) {
                resolve(item.itemId);
            } else if (item.getItemIdAsync) {
                item.getItemIdAsync((result: Office.AsyncResult<string>) => {
                    if (result.status == Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        reject(result.error);
                    }
                });
            } else {
                reject("Cannot get either InternetMessageId, itemId property or getItemIdAsync() method.")
            }
        });
    }
    private static async _getItemBody(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<string> {
        return new Promise((resolve, reject) => {
            item.body.getAsync(Office.CoercionType.Html, (result: Office.AsyncResult<string>) => {
                if (result.status == Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(result.error);
                }
            });
        });
    }
    private static async _getItemSubject(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<string> {
        return new Promise((resolve, reject) => {
            if (typeof item.subject === "string") {
                resolve(item.subject);
            } else {
                const subcallback = item.subject as Office.Subject;
                if (subcallback != null) {
                    subcallback.getAsync((result: Office.AsyncResult<string>) => {
                        if (result.status == Office.AsyncResultStatus.Succeeded)
                            resolve(result.value);
                        else
                            reject(result.error);
                    });
                }
            }
        });
    }
    private static async _getItemTos(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<string[]> {
        return new Promise((resolve, reject) => {
            if (isArray(item.to)) {
                resolve(item.to.map(e => e.emailAddress));
            } else {
                const subcallback = item.to as Office.Recipients;
                if (subcallback != null) {
                    subcallback.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
                        if (result.status == Office.AsyncResultStatus.Succeeded)
                            resolve(result.value.map(e => e.emailAddress));
                        else
                            reject(result.error);
                    });
                }
            }
        });
    }
    private static _getItemCategories(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<string[]> {
        return new Promise((resolve, reject) => {
            const subcallback = item.categories as Office.Categories;
            if (subcallback != null) {
                subcallback.getAsync((result: Office.AsyncResult<Office.CategoryDetails[]>) => {
                    if (result.status == Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value ? result.value.map(e => e.displayName) : null);
                    } else
                        reject(result.error);
                });
            }
        });
    }
    private static _getItemFrom(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): string {
        return item.from.emailAddress;
    }
    private static async _getItemStart(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<moment.Moment> {
        return new Promise((resolve, reject) => {
            if (item.start instanceof Date) {
                resolve(moment(item.start));
            } else {
                const subcallback = item.start as Office.Time
                if (subcallback != null) {
                    subcallback.getAsync((result: Office.AsyncResult<Date>) => {
                        if (result.status == Office.AsyncResultStatus.Succeeded)
                            resolve(moment(result.value));
                        else
                            reject(result.error);
                    });
                }
            }
        });
    }
    private static async _getItemEnd(item: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead): Promise<moment.Moment> {
        return new Promise((resolve, reject) => {
            if (item.end instanceof Date) {
                resolve(moment(item.end));
            } else {
                const subcallback = item.end as Office.Time
                if (subcallback != null) {
                    subcallback.getAsync((result: Office.AsyncResult<Date>) => {
                        if (result.status == Office.AsyncResultStatus.Succeeded)
                            resolve(moment(result.value));
                        else
                            reject(result.error);
                    });
                }
            }
        });
    }

}
export class OutlookItemJSON {
    public ItemId: string;
    public ItemType: string;
    //private _CATEGORIES = ["To Do", "Doing", "Done"];
    constructor(itemId: string, itemType: string) {
        this.ItemId = itemId;
        this.ItemType = itemType;
    }
    public toJson(): string {
        return JSON.stringify(this);
    }
    public async setCategory(category: string, isDialog: boolean = false): Promise<string> {
        let ret = "";
        try {
            if (isDialog) {
                const dialogMessage = new DialogMessageUpdateCategory(MessageActionType.UpdateCategory, this, category);
                Office.context.ui.messageParent(dialogMessage.toJson());
                console.log(dialogMessage);
            } else {
                // Check if master category includes those task's state category
                info(`Starting to apply category.`);
                const s = moment();
                if (await this._isMasterCategoriesReady()) {
                    await this._applyCategory(category);
                }
                const e = moment();
                info(`Updating category duration: ${e.diff(s, 'seconds')} seconds.`);
                console.log(`Updating category duration: ${e.diff(s, 'seconds')} seconds.`);
            }

        } catch (error) {
            console.log("setCategory:", error);
            ret = error.toString();
        }
        return ret;
    }
    private async _applyCategory(category): Promise<void> {
        return new Promise(async (resolve, reject) => {
            const owsToken = await OwsTokenHelper.getOwsToken();
            if (owsToken != null && owsToken != "") {
                // var ewsId = Office.context.mailbox.item.itemId;
                //var token = result.value;
                var getMessageUrl = "";
                if (this.ItemId.indexOf("@") != -1) {//$select=Id,Categories&
                    getMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages?$filter=InternetMessageId eq '${this.ItemId}'`;
                } else {
                    getMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${this.ItemId}`;//?$select=Id
                }

                //console.log(getMessageUrl);
                //console.log(token);
                //info(`_applyCategory-Updating category`);

                const idResponse = await fetch(getMessageUrl, {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${owsToken}`
                    }
                });
                const idResponseJsonObj = await idResponse.json();

                const msgObj = isArray(idResponseJsonObj.value) ? idResponseJsonObj.value[0] : idResponseJsonObj.value;
                const itemId = msgObj["Id"];
                let existingCategories: string[] = [category];
                if (msgObj["Categories"]) {
                    forEach(msgObj["Categories"], c => {
                        const stateCategory = Common.CATEGORIES.find(ec => ec.toLowerCase() == c.toLowerCase())
                        if (stateCategory == null) {
                            existingCategories.push(c);
                        }
                    });
                }
                //info(`_applyCategory-getting existing category`);

                const updateMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${itemId}`;
                // console.log("updateMessageUrl", updateMessageUrl);
                // console.log("token:", token);
                await fetch(updateMessageUrl, {
                    method: 'PATCH',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${owsToken}`
                    },
                    body: JSON.stringify({
                        Categories: existingCategories
                    })
                });
                resolve();
            } else {
                reject("Token is empty");
            }
        });
    }
    private async _isMasterCategoriesReady(): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            //info(`Checking if master categories are ready.`);
            Office.context.mailbox.masterCategories.getAsync(result => {
                if (result.status == Office.AsyncResultStatus.Succeeded && result.value) {
                    const existCategories = result.value.filter(item => {
                        return Common.CATEGORIES.find(c => c.toLowerCase() == item["displayName"].toLowerCase()) != null;
                    });
                    //info(`master categories are ready.`);
                    resolve(existCategories && existCategories.length == 3);
                } else {
                    resolve(false);
                }
            });

        });
    }
    public async popupForm(isDialog: boolean = false): Promise<void> {
        if (isDialog) {
            const dialogMessage = new DialogMessagePopupEmail(MessageActionType.PopupEmail, this);
            Office.context.ui.messageParent(dialogMessage.toJson());
            console.log(dialogMessage);
        } else {
            switch (this.ItemType) {
                case "message":
                    if (this.ItemId.indexOf("@") != -1) {
                        // display email based on the internetMessageId.
                        await this._displayEmail(this.ItemId);
                    } else {
                        // display email based on the ItemId.
                        Office.context.mailbox.displayMessageForm(this.ItemId);
                    }
                    break;
                case "appointment":
                    Office.context.mailbox.displayAppointmentForm(this.ItemId);
                    break;
                default:
                    break;
            }
        }
    }
    private async _displayEmail(internetMessageId: string): Promise<void> {
        return new Promise(async (resolve, reject) => {
            var owsToken = await OwsTokenHelper.getOwsToken();
            var getMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages?$select=Id&$filter=InternetMessageId eq '${internetMessageId}'`;

            const idResponse = await fetch(getMessageUrl, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${owsToken}`
                }
            });
            const idResponseJsonObj = await idResponse.json();
            const itemId = idResponseJsonObj.value[0]["Id"];
            Office.context.mailbox.displayMessageForm(itemId);
            resolve();

        });
    }

    public static createInstance(json: string): OutlookItemJSON {
        let ret: OutlookItemJSON = OutlookItemJSON.createTaskInstance();
        if (json != null && json != "") {
            if (json.indexOf("{") == -1) {
                ret = new OutlookItemJSON(json, "message");
            } else {
                const temp = JSON.parse(json) as OutlookItemJSON;
                ret = new OutlookItemJSON(temp.ItemId, temp.ItemType);
            }
        }
        return ret;
    }
    public static createTaskInstance(): OutlookItemJSON {
        return new OutlookItemJSON("", "task");
    }
}