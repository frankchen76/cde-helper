import * as React from "react";
import { useState, useEffect } from "react";
import Progress from "./Progress";
import Header from "./Header";
import TasksView from "./TasksView";
import TaskItemView from "./TaskItemView";
import {
    HashRouter,
    Redirect,
    Switch,
    Route
} from "react-router-dom";

import { IServiceContext, ServiceContext, Setting } from "../services/SettingService";
import { IPackageInfo } from "../services/IPackageInfo";
import { IconButton, Spinner, SpinnerSize, Stack, TextField } from "@fluentui/react";
import { IssueService } from "../services/IssueService";
// import { withAITracking } from '@microsoft/applicationinsights-react-js';
// import { reactPlugin, appInsights } from '../services/AppInsights';
import IssuesView from "./IssuesView";
import IssueItemView from "./IssueItemView";
import TasksReport from "./TasksReport";
import SettingsView from "./SettingsView";
import { TaskService } from "../services/TaskService";
import { QueryService } from "../services/QueryService";
import { AreaService } from "../services/AreaService";
import { TagService } from "../services/TagService";
import { ReportService } from "../services/ReportService";
import { OutlookItem } from "../services/OutlookItem";
import TasksViewByArea from "./TasksViewByArea";
import { HostInfo, OriginType } from "../services/HostInfo";

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

const App = (props: AppProps) => {
    const [currentMailItem, setCurrentMailItem] =
        useState<Office.Item
            & Office.ItemCompose
            & Office.ItemRead
            & Office.Message
            & Office.MessageCompose
            & Office.MessageRead
            & Office.Appointment
            & Office.AppointmentCompose
            & Office.AppointmentRead>(Office.context?.mailbox?.item);
    //const [settingService, setSettingService] = useState<SettingService>();
    const [serviceContext, setServiceContext] = useState<IServiceContext>();
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [outlookItem, setOutlookItem] = useState<OutlookItem>();

    // https://localhost:3443/taskpane.html?_host_Info=Outlook$Win32$16.02$en-US$$$$0#/  from outlook taskpane
    // https://localhost:3443/taskpane.html?_host_Info=Outlook$Win32$16.02$en-US$telemetry$isDialog$$0#/  from outlook dialog
    // const currentUrl = new URL(window.location.href);
    // const hostInfo = currentUrl.searchParams.get("_host_Info");
    // const [isDialog, setIsDialog] = useState<boolean>(hostInfo && hostInfo.indexOf("isDialog") != -1);
    const hostInfo = new HostInfo();

    const { title, isOfficeInitialized } = props;
    const containerStyle = {
        maxWidth: "800px",
        margin: "auto"
    };
    const packageInfo: IPackageInfo = require("../../../package.json");

    // Load SettingService2
    useEffect(() => {
        setIsLoading(true);
        const loadSettingService = async () => {
            const setting = await Setting.getSetting();
            const taskService = new TaskService();
            const issueService = new IssueService();
            taskService._issueService = issueService;
            let selectedOutlookItem: OutlookItem = null;
            if (Office.context?.mailbox?.item) {
                selectedOutlookItem = await OutlookItem.createInstance(Office.context.mailbox.item);
                setOutlookItem(selectedOutlookItem);
            }
            const context = {
                taskService: taskService,
                issueService: issueService,
                queryService: new QueryService(),
                areaService: new AreaService(),
                tagService: new TagService(),
                reportService: new ReportService(),
                setting: setting,
                hostInfo: hostInfo,
                selectedOutlookItem: selectedOutlookItem,
                onSettingUpdate: onSettingUpdate
            };

            // appInsights.setAuthenticatedUserContext(setting.upn, setting.upn);
            setServiceContext(context);
            setIsLoading(false);
        };

        loadSettingService();
    }, []);

    // Load mainitem
    useEffect(() => {
        // add handler if it's not in dialog
        if (hostInfo.Origin != OriginType.OutlookDialog) {
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, (): void => {
                if (Office.context.mailbox.item) {
                    OutlookItem.createInstance(Office.context.mailbox.item).then(newItem => {
                        if (newItem) {
                            setOutlookItem(newItem);
                            // update service context to include the latestest outlook item
                            setServiceContext(context => ({ ...context, selectedOutlookItem: newItem }));
                        }
                    });
                }
            });
        }
    }, []);

    const onSettingUpdate = (val: Setting) => {
        setServiceContext(context => ({ ...context, setting: val }));
    };

    // Rendering NOTE: dir="ltr" is required
    if (isLoading) {
        return (<Spinner title="Loading add-in..." />);
    } else {
        return (<HashRouter>
            <div className="ms-Grid" style={containerStyle} dir="ltr">
                <ServiceContext.Provider value={serviceContext} >
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style={{ paddingLeft: 0 }}>
                            <Header />
                        </div>
                    </div>
                    <div className="ms-Grid-row" style={{ "marginTop": "44px" }}>
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style={{ minHeight: 680 }}>
                            <Switch>
                                {/* <Route path="/taskitem/:settingItemId/:id" render={(routeProps) => <TaskItemView outlookItem={outlookItem} settingItemId={routeProps["match"]["params"]["settingItemId"]} id={+routeProps["match"]["params"]["id"]} />} /> */}
                                <Route path="/taskitem/:settingItemId/:id">
                                    <TaskItemView outlookItem={outlookItem} />
                                </Route>
                                <Redirect from="/redirecttaskitem/:settingItemId/:id" to="/taskitem/:settingItemId/:id" />
                                <Route path="/issuesview" render={(routeProps) => <IssuesView routeProps={routeProps} />} />
                                {/* <Route path="/issueitem/:settingItemId/:id" render={(routeProps) => <IssueItemView routeProps={routeProps} />} /> */}
                                <Route path="/issueitem/:settingItemId/:id">
                                    <IssueItemView />
                                </Route>
                                <Redirect from="/redirectissueitem/:settingItemId/:id" to="/issueitem/:settingItemId/:id" />
                                <Route path="/tasksreport" render={(routeProps) => <TasksReport routeProps={routeProps} />} />
                                <Route path="/setting" render={(routeProps) => <SettingsView routeProps={routeProps} />} />
                                <Route path="/tasksviewbyarea/:settingItemId/:areaId">
                                    <TasksViewByArea />
                                </Route>
                                <Route path="/" render={(routeProps) => <TasksView routeProps={routeProps} />} />
                            </Switch>
                        </div>
                    </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style={{ paddingLeft: 0 }}>
                            <div className="footerbar">
                                <Stack horizontal horizontalAlign="space-between">
                                    <div>
                                        <span>{`Setting Name: ${serviceContext.setting.DefaultSetting && serviceContext.setting.DefaultSetting.name}`}</span>
                                    </div>
                                    <div>{`v${packageInfo.version}`}</div>
                                </Stack>
                            </div>
                        </div>
                    </div>
                </ServiceContext.Provider>
                {/* </SettingServiceContext.Provider> */}
            </div>
        </HashRouter>)
    }

};
//export default withAITracking(reactPlugin, App);
export default App;
