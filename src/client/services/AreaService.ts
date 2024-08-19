import { find } from "lodash";
import { Area, AreaCollection } from "./Area";
import { BaseService } from "./BaseService"
import { ISettingItem } from "./SettingService";
import moment from "moment";
import { OutlookItem } from "./OutlookItem";

export interface IAreaService {
    getAreasByBaseUrl(baseUrl: string): Promise<AreaCollection>;
    getAreas(settingItem: ISettingItem): Promise<AreaCollection>;
    getAreas(settingItem: ISettingItem, filterBySetting: boolean): Promise<AreaCollection>;
    getCurrentIterationId(settingItem: ISettingItem): Promise<number>;
    getAreaByOutlookItem(settingItem: ISettingItem, areas: AreaCollection, outlookItem: OutlookItem): Area;
    getAreaById(settingItem: ISettingItem, areaId: number): Promise<Area>;
}
export class AreaService extends BaseService {

    public async getAreasByBaseUrl(baseUrl: string): Promise<AreaCollection> {

        let ret = null;
        let url = `${baseUrl}/_apis/wit/classificationnodes?$depth=2&api-version=6.0`;
        const response = await this._httpClientService.get(url);
        if (response != null) {
            ret = AreaCollection.createAreasFromResponse(null, response, false);
        }
        return ret;
    }
    public async getAreas(settingItem: ISettingItem, filterBySetting: boolean = true): Promise<AreaCollection> {

        let ret = null;
        let url = `${settingItem.baseUrl}/_apis/wit/classificationnodes?$depth=2&api-version=6.0`;
        const response = await this._httpClientService.get(url);
        if (response != null) {
            ret = AreaCollection.createAreasFromResponse(settingItem, response, filterBySetting);
        }
        return ret;
    }
    public async getCurrentIterationId(settingItem: ISettingItem): Promise<number> {
        let ret = 0;
        let url = `${settingItem.baseUrl}/_apis/wit/classificationnodes?$depth=2&api-version=6.0`;
        const response = await this._httpClientService.get(url);
        if (response != null && response.value != null) {
            const areaRows = find(response["value"], c => c["structureType"] == "iteration");
            if (areaRows) {
                const currentIternation = areaRows["children"].find(row => {
                    let ret = false;
                    if (row["attributes"] && row["attributes"]["startDate"] && row["attributes"]["finishDate"]) {
                        const start = moment(row["attributes"]["startDate"], "YYYY-MM-DD");
                        const end = moment(row["attributes"]["finishDate"], "YYYY-MM-DD");
                        const firstDayOfThisWeek = moment().day(0).startOf("day");
                        const lastDayOfThisWeek = moment().day(6).startOf("day");
                        //ret = today.isSameOrBefore(end, "day") && today.isSameOrAfter(start, "day");
                        ret = start >= firstDayOfThisWeek && end <= lastDayOfThisWeek;
                    }
                    return ret;
                });
                ret = currentIternation ? currentIternation["id"] : 0;
            }
        }
        return ret;
    }
    public getAreaByOutlookItem(settingItem: ISettingItem, areas: AreaCollection, outlookItem: OutlookItem): Area {
        let ret: Area = areas.items[0]; // return the first area
        if (settingItem.areas) {
            for (const settingAreaItem of settingItem.areas) {
                if (settingAreaItem.enabled) {
                    const existArea = areas.getAreaById(settingAreaItem.areaId);
                    // if outlookitem's to emails include the email domains. 
                    if (outlookItem.areaExistInCategories(existArea.Name) ||
                        outlookItem.areaExistInSubject(existArea.Name) ||
                        (settingAreaItem.emailDomains && outlookItem.emailDomainExists(settingAreaItem.emailDomains))) {

                        ret = existArea;//areas.getAreaById(settingAreaItem.areaId);
                        break;
                    }
                }
            }
        }
        return ret;
    }
    public async getAreaById(settingItem: ISettingItem, areaId: number): Promise<Area> {
        const allAreas = await this.getAreas(settingItem);
        return allAreas.getAreaById(areaId);
    }

}