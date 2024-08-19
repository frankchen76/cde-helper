import { IDropdownOption, SelectableOptionMenuItemType } from "@fluentui/react";

export class Query {
    // private _id: string;
    // private _queryName: string;
    // private _isHeader: boolean;
    // private _childen = new QueryCollection(null);

    // public get Id(): string { return this._id; }
    // public get QueryName(): string { return this._queryName; }
    // public get IsHeader(): boolean { return this._isHeader; }
    // public get Children(): QueryCollection { return this.Children; }
    public id: string;
    public queryName: string;
    public isHeader: boolean;
    public childen: QueryCollection;

    // constructor(responseRow: any) {
    //     this.id = responseRow["id"];
    //     this.queryName = responseRow["name"];
    //     this.isHeader = responseRow["isFolder"];

    //     if (responseRow["hasChildren"]) {
    //         responseRow["children"].forEach(queryResponse => {
    //             this.childen.push(new Query(queryResponse));
    //         });
    //     }
    // }
    public static createQueryFromResponse(response: any): Query {
        let ret = new Query();
        ret.id = response["id"];
        ret.queryName = response["name"];

        return ret;
    }
}

export class QueryCollection {
    constructor(public items: Query[]) {

    }
    public static createQueriesFromResponse(response: any): QueryCollection {
        let ret = null;
        if (response && response.value) {
            ret = new QueryCollection([]);
            response.value.forEach(itemResponse => {
                let qHeader = Query.createQueryFromResponse(itemResponse);
                if (itemResponse["children"]) {
                    qHeader.childen = new QueryCollection([]);
                    itemResponse["children"].forEach(subitemResponse => {
                        let qItem = Query.createQueryFromResponse(subitemResponse);
                        qHeader.childen.items.push(qItem);
                    });
                }
                ret.items.push(qHeader);
            })
        }
        return ret;
    }
    public toIDropdownOptions(): IDropdownOption[] {
        let ret: IDropdownOption[] = [];
        this.items.forEach(item => {
            ret.push({
                key: item.id,
                text: item.queryName,
                itemType: SelectableOptionMenuItemType.Header
            });

            if (item.childen) {
                item.childen.items.forEach(subItem => {
                    ret.push({
                        key: subItem.id,
                        text: subItem.queryName,
                        itemType: SelectableOptionMenuItemType.Normal
                    });
                });
            }
        });
        return ret;
    }
}