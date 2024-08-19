import { findIndex, groupBy, orderBy, sumBy } from 'lodash';
import moment from "moment";
import { Area } from './Area';
import { Issue } from './Issue';

export class ReportItem {
    public Id: number;
    public Title: string;
    public CreatedDate: moment.Moment;
    public StateChangeDate: moment.Moment;
    public State: string;
    public Area: Area;
    public Issue: Issue;
    public CompletedWork: number;
    public TodayHours: number;

    constructor(id: number,
        title: string,
        state: string,
        createdDate: string,
        stateChangeDate: string,
        areaId: number,
        areaName: string,
        areaPath: string,
        completedWork: number) {
        this.Id = id;
        this.Title = title;
        this.State = state;
        this.CreatedDate = moment.utc(createdDate).local();
        this.StateChangeDate = moment.utc(stateChangeDate).local();
        this.Area = new Area(areaId, areaName, areaPath, 0);
        this.CompletedWork = completedWork;
    }
    public static createInstanceFromJSON(responseRow: any): ReportItem {
        return new ReportItem(responseRow["id"],
            responseRow["fields"]["System.Title"],
            responseRow["fields"]["System.State"],
            responseRow["fields"]["System.CreatedDate"],
            responseRow["fields"]["Microsoft.VSTS.Common.StateChangeDate"],
            responseRow["fields"]["System.AreaId"],
            responseRow["fields"]["System.NodeName"],
            responseRow["fields"]["System.AreaPath"],
            responseRow["fields"]["Microsoft.VSTS.Scheduling.CompletedWork"]
        );
    }

    public IsCreatedSameAsClosed(): boolean {
        return this.State == "Done" && this.CreatedDate.isSame(this.StateChangeDate, "day");
    }
}

export class ReportItemGroup {
    public get GroupName() { return this._groupName; }
    public get ReportItems() { return this._reportItems; }

    constructor(private _groupName: string, private _reportItems: ReportItem[]) {

    }
    public TotalHours(): number {
        return sumBy(this._reportItems, "TodayHours")
    }
}
export class ReportItemGroupCollection {
    public get Groups() { return this._group; }
    constructor(private _group: ReportItemGroup[]) {

    }
    public TotalHours(): number {
        return sumBy(this._group, g => g.TotalHours());
    }

}

export class ReportItemCollection {
    private _items: ReportItem[];
    public get Items() { return this._items; }

    constructor() {
        this._items = [];
    }
    public sortByArea(): void {
        this._items = orderBy(this._items, "Area.Name", "desc")
    }
    public groupByIssueArea(): ReportItemGroupCollection {
        let groups: ReportItemGroup[] = [];
        let areaIssues = groupBy(this._items, (item: ReportItem): string => {
            //return `${item.Area.Name}-${item.Issue.title}`;
            return `${item.Area.Name}-${item.Issue.axisCode ? item.Issue.axisCode : item.Issue.title}`;
        });
        for (let name in areaIssues) {
            groups.push(new ReportItemGroup(name, areaIssues[name]));
        }
        return new ReportItemGroupCollection(groups);
    }
    public addReportItems(newItems: ReportItem[]) {
        this._items = this._items.concat(newItems);
    }
}

export class HistoryItem {
    public ChangedDate: moment.Moment;
    public CompletedWork: number
    constructor(changedDate: string, completedWork: number, isUtc: boolean = true) {
        this.ChangedDate = isUtc ? moment.utc(changedDate).local() : moment(changedDate);
        this.CompletedWork = completedWork;
    }
    public static createInstanceFromJSON(responseRow: any): HistoryItem {
        return new HistoryItem(responseRow["fields"]["System.ChangedDate"],
            responseRow["fields"]["Microsoft.VSTS.Scheduling.CompletedWork"]);
    }
}
export class HistoryItemCollection {
    private _histories: HistoryItem[];
    public get Histories(): HistoryItem[] { return this._histories; }

    constructor() {
        this._histories = [];
    }

    public static createInstanceFromJSON(response: any): HistoryItemCollection {
        let ret = new HistoryItemCollection();
        response["value"].forEach(row => {
            ret.Histories.push(HistoryItem.createInstanceFromJSON(row));
        });
        return ret;
    }

    public getSelectedDateHour(date: moment.Moment = moment()): number {
        let ret = 0;
        const dateHours = groupBy(this._histories, (item: HistoryItem) => {
            return item.ChangedDate.format('YYYY-MM-DD');
        });
        if (dateHours) {
            let groupByDateResult: HistoryItem[] = [];
            for (let d in dateHours) {
                const lastHour = dateHours[d][dateHours[d].length - 1];
                groupByDateResult.push(new HistoryItem(d, lastHour.CompletedWork, false));
            }
            const todayIndex = findIndex(groupByDateResult, d => d.ChangedDate.isSame(date, "day"));
            if (todayIndex != -1) {
                ret = todayIndex == 0 ? groupByDateResult[todayIndex].CompletedWork :
                    groupByDateResult[todayIndex].CompletedWork - groupByDateResult[todayIndex - 1].CompletedWork;
            }
        }
        return ret;
    }
}