import * as _ from "lodash";
import { IDropdownOption } from "@fluentui/react";
import { ISettingAreaItem, ISettingItem } from "./SettingService";
import { OutlookItem } from "./OutlookItem";

export class Area {
    public get Id(): number { return this._id; }
    public get Name(): string { return this._name; }
    public get Path(): string { return this._path; }
    public get WiqlPath(): string {
        const frags = this._path.split("\\");
        let ret = [frags[1], frags[3]];
        return ret.join("\\");
    }
    public get ParentId(): number { return this._parentId; }

    constructor(private _id: number,
        private _name: string,
        private _path: string,
        private _parentId: number) {
    }
    public toIDropdownOption(): IDropdownOption {
        return {
            key: this._id,
            text: this._name
        }
    }
    public static createInstanceFromJSON(row: any, parentId: number): Area {
        const id = row["id"];
        const name = row["name"];
        const path = row["path"];
        return new Area(id, name, path, parentId);
    }
}

export class AreaCollection {
    constructor(public items: Area[]) {

    }
    public static createAreasFromResponse(settingItem: ISettingItem, response: any, filterBySetting: boolean = true): AreaCollection {
        let ret: AreaCollection = null;
        if (response && response.value) {
            ret = new AreaCollection([]);
            const areaRows = _.find(response["value"], { "structureType": "area" });
            if (areaRows) {
                areaRows["children"].forEach(row => {
                    let settingAreaItem: ISettingAreaItem;
                    // skip this if filterBySetting=false. this is for SettingView.
                    if (filterBySetting) {
                        settingAreaItem = _.find(settingItem.areas, { "areaId": row["id"] }) as ISettingAreaItem;
                    }
                    if (settingAreaItem == null || (settingAreaItem && settingAreaItem.enabled)) {
                        const newArea = Area.createInstanceFromJSON(row, +areaRows["id"]);
                        ret.items.push(newArea);
                    }
                });
                //this._items = orderBy(this._items, ["Name"], ["asc"]);
                ret.sort();
            }
        }
        return ret;
    }
    public sort() {
        //this.items = orderBy(this.items, ["Name"], ["asc"]);
        this.items = this.items.sort((a, b) => {
            return a.Name.toLowerCase().localeCompare(b.Name.toLowerCase());
        });
    }
    public toIDropdownOption(settingItem: ISettingItem): IDropdownOption[] {
        let ret: IDropdownOption[] = [];
        if (this.items) {
            // this.items.forEach(item => {
            //     const areaSetting = settingItem.areas.find(a => a.areaId == item.Id);
            //     if (areaSetting && areaSetting.enabled) {
            //         ret.push(item.toIDropdownOption());
            //     } else if (!areaSetting) {
            //         ret.push(item.toIDropdownOption());
            //     }
            // });
            ret = this.items.map(t => t.toIDropdownOption());
        }
        return ret;
    }
    public getAreaById(id: number): Area {
        return _.find(this.items, a => a.Id == id);
    }
}