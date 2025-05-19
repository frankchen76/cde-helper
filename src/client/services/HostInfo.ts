export enum OriginType {
    Web = "web",
    OutlookDialog = "outlookdialog",
    OutlookTaskPane = "outlooktaskpane",
    MSTeams = "msteams",
}
export class HostInfo {
    private _hostUrl: URL;
    private static _qsOrigin = "origin";
    private _origin: OriginType = OriginType.Web;
    public get Origin(): OriginType {
        return this._origin;
    }
    constructor() {
        this._hostUrl = new URL(window.location.href);
        // https://localhost:3443/taskpane.html?_host_Info=Outlook$Win32$16.02$en-US$$$$0#/  from outlook taskpane
        // https://localhost:3443/taskpane.html?_host_Info=Outlook$Win32$16.02$en-US$telemetry$isDialog$$0#/  from outlook dialog
        const hostInfo = this._hostUrl.searchParams.get("_host_Info");
        // URL will contain "origin" query string defined in OriginType enum to determine the host type
        const qsOrigin = this._hostUrl.searchParams.get(HostInfo._qsOrigin);
        if (Object.values(OriginType).some((col: string) => col === qsOrigin)) {
            this._origin = <OriginType>qsOrigin;
            // if origin is outlook taskpane, check if it is a dialog
            if (this._origin == OriginType.OutlookTaskPane && hostInfo && hostInfo.indexOf("isDialog") != -1) {
                this._origin = OriginType.OutlookDialog;
            }
        }
    }
    public get IsOutlookTaskPane(): boolean {
        return this._origin === OriginType.OutlookTaskPane;
    }

    public static GenerateQSOrigin(origin: OriginType): string {
        return `${HostInfo._qsOrigin}=${origin}`;
    }
    public static getHostInfo(): HostInfo {
        return new HostInfo();
    }
    // public static getOrigin(): OriginType {
    //     const hostUrl = new URL(window.location.href);
    //     const qsOrigin = hostUrl.searchParams.get(HostInfo._qsOrigin);
    //     var ret = OriginType.Web;
    //     if (Object.values(OriginType).some((col: string) => col === qsOrigin)) {
    //         ret = <OriginType>qsOrigin;
    //     }
    //     return ret;
    // }
}