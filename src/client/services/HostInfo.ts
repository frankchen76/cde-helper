export enum OriginType {
    OutlookDialog = "outlookdialog",
    OutlookTaskPane = "outlooktaskpane",
    MSTeams = "msteams",
}
export class HostInfo {
    private _hostUrl: URL;
    private static _qsOrigin = "origin";
    private _origin: OriginType;
    public get Origin(): OriginType {
        return this._origin;
    }
    constructor() {
        this._hostUrl = new URL(window.location.href);
        // https://localhost:3443/taskpane.html?_host_Info=Outlook$Win32$16.02$en-US$$$$0#/  from outlook taskpane
        // https://localhost:3443/taskpane.html?_host_Info=Outlook$Win32$16.02$en-US$telemetry$isDialog$$0#/  from outlook dialog
        const hostInfo = this._hostUrl.searchParams.get("_host_Info");
        const qsOrigin = this._hostUrl.searchParams.get(HostInfo._qsOrigin);
        if (Object.values(OriginType).some((col: string) => col === qsOrigin)) {
            this._origin = <OriginType>qsOrigin;
        }
    }
    public static GenerateQSOrigin(origin: OriginType): string {
        return `${HostInfo._qsOrigin}=${origin}`;
    }
}