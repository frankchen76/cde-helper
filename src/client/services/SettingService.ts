import jwtDecode, { JwtPayload } from "jwt-decode";
import { ITaskService, TaskService } from "./TaskService";
import { IIssueService, IssueService } from "./IssueService";
import { AreaService, IAreaService } from "./AreaService";
import { IQueryService, QueryService } from "./QueryService";
import { IReportService, ReportService } from "./ReportService";
import { ITagService, TagService } from "./TagService";
import { IDropdownOption } from "@fluentui/react";
import { createContext, useContext } from "react";
import { ProfileService } from "./ProfileService";
import { HostInfo } from "./HostInfo";
import { OutlookItem } from "./OutlookItem";
import { info } from "./log";
import { err } from "../../services/log";

export interface ISettingItem {
    id: string;
    name: string;
    tasksQueryId: string;
    issuesQueryId: string;
    reportQueryId: string;
    baseUrl: string;
    //email: string;
    areas?: ISettingAreaItem[];
}

export interface ISettingAreaItem {
    areaId: number;
    areaName?: string;
    enabled: boolean;
    emailDomains: string[];
}

export class Setting {
    public static SETTING_NAME = "AzureDevOpsAddinSetting";
    public static DEFAULT_SETTINGNAME = "CDE2";
    constructor(public items: ISettingItem[],
        public upn: string,
        public defaultSettingId: string,
        public apiKey: string = "") {

    }
    public get DefaultSetting(): ISettingItem {
        //return this.items.find(s => s.id == this.defaultSettingId);
        return this.items && Array.isArray(this.items) ? this.items.find(s => s.id == this.defaultSettingId) : null;
    }
    public static async getSetting(): Promise<Setting> {
        let localSetting: Setting = Office.context.roamingSettings.get(Setting.SETTING_NAME) as Setting;
        // TODO: force to load from web
        //localSetting = null;
        let ret: Setting = null;
        info("getSetting-localSetting", localSetting);
        if (localSetting) {
            // read from local if available
            ret = new Setting(localSetting.items, localSetting.upn, localSetting.defaultSettingId, localSetting.apiKey);
        } else {
            // load from express.js
            // get current UPN
            //const upn = await Setting.getUpn();
            const upn = await ProfileService.getUpn();
            let url = `${location.protocol}//${location.host}/api/getSettings/${upn}`;
            info(`url: ${url}`);
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'cache': "no-store"
                }
            });
            const remoteSetting: ISettingItem[] = await response.json() as ISettingItem[];
            if (remoteSetting) {
                ret = new Setting(remoteSetting, upn, Setting.DEFAULT_SETTINGNAME);
            }
            info("getSetting-ret", ret);
            await ret.saveSetting();
        }

        return ret;
    }
    public async saveSetting(): Promise<void> {
        return new Promise((resolve, reject) => {
            Office.context.roamingSettings.set(Setting.SETTING_NAME, this);
            Office.context.roamingSettings.saveAsync(result => {
                if (result.status == Office.AsyncResultStatus.Succeeded) {
                    info("saveSetting:successed");
                    resolve();
                } else {
                    err("saveSetting:failed", result.error);
                    reject(result.error);
                }
            });

        });
    }
    public static async removeSetting(): Promise<boolean> {
        return new Promise<boolean>((resolve, reject) => {
            Office.context.roamingSettings.remove(Setting.SETTING_NAME);
            Office.context.roamingSettings.saveAsync((result: Office.AsyncResult<void>) => {
                //update back BaseUrl
                resolve(true);
                console.log(`Setting: ${Setting.SETTING_NAME} removed. ${result.status}`);
            });
        });
    }

    public getSettingItemById(id: string): ISettingItem {
        let ret = this.items.find(i => i.id == id);
        if (!ret) {
            ret = this.DefaultSetting;
        }
        return ret;
    }
    public toIDropdownOption(): IDropdownOption[] {
        let ret: IDropdownOption[] = [];
        if (this.items) {
            ret = this.items.map(t => {
                return {
                    key: t.id,
                    text: t.name
                };
            });
        }
        return ret;
    }
};
export interface IServiceContext {
    setting: Setting;
    hostInfo: HostInfo;
    taskService: ITaskService;
    issueService: IIssueService;
    queryService: IQueryService;
    areaService: IAreaService;
    tagService: ITagService;
    reportService: IReportService;
    selectedOutlookItem?: OutlookItem
    onSettingUpdate: (val: Setting) => void;
};
export const ServiceContext = createContext<IServiceContext>(null);


// export class SettingService {
//     private _setting: ISetting;
//     private _upn: string;
//     private _token: AzureDevOpsToken;
//     private SETTING_NAME = "AzureDevOpsAddin";
//     private _onSettingServiceChangedCallback: (settingService: SettingService) => void;

//     public get Setting(): ISetting { return this._setting; }
//     public get DefaultSettingItem(): ISettingItem {
//         return this._setting.items.find(item => item.id == this._setting.defaultSettingItemId)
//     }
//     public get SettingItemOptions(): IDropdownOption[] {
//         let ret: IDropdownOption[] = [];
//         if (this._setting.items) {
//             ret = this._setting.items.map(item => {
//                 return {
//                     key: item.name,
//                     text: item.name
//                 }
//             });
//         }
//         return ret;
//     }
//     public get Token(): AzureDevOpsToken { return this._token; }
//     public set Token(val: AzureDevOpsToken) { this._token = val; }

//     constructor(onSettingServiceUpdateCallback: (settingService: SettingService) => void) {
//         this._token = AzureDevOpsToken.createTokenFromHTML();
//         this._onSettingServiceChangedCallback = onSettingServiceUpdateCallback;
//         let addinSetting: ISetting = Office.context.roamingSettings.get(this.SETTING_NAME) as ISetting;
//         //Office.context.u

//         if (!addinSetting) {
//             const defaultNewItem = this.createSettingItem();
//             addinSetting = {
//                 isFirstTime: true,
//                 defaultSettingItemId: defaultNewItem.id,
//                 items: [defaultNewItem]
//             }
//         } else {
//             if (addinSetting.items == undefined) {
//                 //upgrade from old to new
//                 const settingItems: ISettingItem[] = [{
//                     id: Guid.create().toString(),
//                     name: "Default",
//                     tasksQueryId: addinSetting["tasksQueryId"],
//                     issuesQueryId: addinSetting["issuesQueryId"],
//                     reportQueryId: addinSetting["reportQueryId"],
//                     baseUrl: addinSetting["baseUrl"],
//                     email: addinSetting["email"],
//                 }];
//                 addinSetting = {
//                     defaultSettingItemId: settingItems[0].id,
//                     isFirstTime: false,
//                     items: settingItems
//                 }
//             }
//             addinSetting.isFirstTime = false;
//             const defaultSettingItem = addinSetting.items.find(item => item.id == addinSetting.defaultSettingItemId);
//         }
//         this._setting = addinSetting;

//     }
//     public createSettingItem(): ISettingItem {
//         return {
//             id: Guid.create().toString(),
//             name: "<New>",
//             tasksQueryId: "",
//             reportQueryId: "",
//             issuesQueryId: "",
//             baseUrl: undefined,
//             email: Office.context.mailbox.userProfile.emailAddress
//         };
//     }

//     public updateSetting(setting: ISetting): Promise<boolean> {
//         this._setting = setting;
//         this._setting.isFirstTime = false;
//         const settingName = this.SETTING_NAME;
//         return new Promise<boolean>((resolve, reject) => {
//             Office.context.roamingSettings.set(settingName, setting);
//             Office.context.roamingSettings.saveAsync((result: Office.AsyncResult<void>) => {
//                 //update back BaseUrl
//                 const defaultSettingItem = setting.items.find(item => item.id == setting.defaultSettingItemId);
//                 if (this._onSettingServiceChangedCallback) {
//                     this._onSettingServiceChangedCallback(this);
//                 }
//                 console.log(`Setting: ${settingName} saved. ${result.status}`);
//                 resolve(true);
//             });
//         });
//     }
//     public removeSetting(): Promise<boolean> {
//         const settingName = this.SETTING_NAME;
//         return new Promise<boolean>((resolve, reject) => {
//             const settingName = this.SETTING_NAME;
//             Office.context.roamingSettings.remove(settingName);
//             Office.context.roamingSettings.saveAsync((result: Office.AsyncResult<void>) => {
//                 //update back BaseUrl
//                 resolve(true);
//                 console.log(`Setting: ${settingName} removed. ${result.status}`);
//             });
//         });
//     }
// }