import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
// import { withAITracking } from "@microsoft/applicationinsights-react-js";
// import { reactPlugin } from "../services/AppInsights";
import { useContext, useEffect, useState } from "react";
import { Common, ExecutingResult } from "../services/Common";
import { ServiceContext } from "../services/SettingService";
import { ReportItemCollection, ReportItemGroup, ReportItemGroupCollection } from "../services/ReportItem";
import { IconButton, MessageBar, MessageBarType, Shimmer, ShimmerElementType, Stack, Toggle } from "@fluentui/react";

export interface ITasksReportProps {
    routeProps: RouteComponentProps;
}
const TasksReport = (props: ITasksReportProps) => {
    const [executingResult, setExecutingResult] = useState<ExecutingResult>(ExecutingResult.createInstance());
    const [reportItes, setReportItems] = useState<ReportItemCollection>();
    const [reportGroups, setReportGroups] = useState<ReportItemGroupCollection>();
    const [showTaskHours, setShowTaskHours] = useState<boolean>(true);
    const serviceContext = useContext(ServiceContext);

    useEffect(() => {
        const loadReports = async () => {
            try {
                const { reportService, setting } = serviceContext;
                setExecutingResult(ExecutingResult.start());
                const allReportItems = await reportService.getReportItems(setting);

                // Log report items to DB
                let dbResult = "";
                if (allReportItems && allReportItems.Items && allReportItems.Items.length > 0) {
                    dbResult = await reportService.logReportItemsToDb(setting.apiKey, allReportItems);
                }

                if (allReportItems) {
                    const groups = allReportItems.groupByIssueArea();
                    setReportGroups(groups);
                }

                setReportItems(allReportItems);
                if (dbResult === "") {
                    setExecutingResult(ExecutingResult.complete(false));
                } else {
                    setExecutingResult(ExecutingResult.complete(true, dbResult, true));
                }

            } catch (error) {
                setExecutingResult(ExecutingResult.complete(true, error, true));
            }
        };

        loadReports();
    }, []);


    const renderTask = (group: ReportItemGroup): any => {
        return (
            <ul id={`ul-${group.GroupName}`} style={{ paddingLeft: "20px" }}>
                {group.ReportItems.map(reportItem => {
                    return (<li id={reportItem.Id.toString()}>
                        {`${reportItem.Title} ${showTaskHours ? "(" + reportItem.TodayHours + ")" : ""}`}
                    </li>);
                })}
            </ul>
        )
    };
    const onCopyHandler = (group: ReportItemGroup): void => {
        const elem = document.createElement('textarea')
        elem.value = group.ReportItems.map(reportItem => reportItem.Title).join("\r\n");

        document.body.append(elem)

        // Select the text and copy to clipboard
        elem.select()
        const success = document.execCommand('copy')
        elem.remove()
    };
    const renderGroup = (groups: ReportItemGroupCollection): any => {
        return (
            <ul id="main" style={{ paddingLeft: "20px" }}>
                {groups.Groups.map(group => {
                    const header = group.GroupName;
                    return (<li id={`li-${group["areaName"]}`}>
                        <span>{`${header} (${group.TotalHours()}h)`}</span>
                        <IconButton iconProps={{ iconName: "Copy" }}
                            title="Copy to clipboard"
                            onClick={onCopyHandler.bind(this, group)} />
                        {renderTask(group)}
                    </li>);
                })}
            </ul>
        )
    };
    //const containerStackTokens: IStackTokens = { childrenGap: 5 };
    const shimmerCategory = [
        { type: ShimmerElementType.gap, width: "15%", height: 30 },
        { type: ShimmerElementType.line, width: "40%", height: 30 },
        { type: ShimmerElementType.gap, width: "45%", height: 30 }
    ];
    const shimmerItem = [
        { type: ShimmerElementType.gap, width: "25%", height: 30 },
        { type: ShimmerElementType.line, width: "75%", height: 30 }];

    const onMessageBarDismiss = () => {
        setExecutingResult(result => ({ ...result, displayMessage: false }));
    };
    const onShowTaskHoursChange = (ev: React.MouseEvent<HTMLElement>, checked?: boolean) => {
        setShowTaskHours(checked);
    };

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>Tasks Report</h2>
                </div>
            </div>
            {executingResult.displayMessage &&
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                        <MessageBar messageBarType={executingResult.isError ? MessageBarType.error : MessageBarType.success}
                            onDismiss={onMessageBarDismiss}
                            isMultiline={false}>{executingResult.message}</MessageBar>
                    </div>
                </div>}
            {executingResult.isRunning ?
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                        <Stack tokens={Common.CONTAINER_STACK_TOKENS}>
                            <Shimmer shimmerElements={shimmerCategory} />
                            <Shimmer shimmerElements={shimmerItem} />
                            <Shimmer shimmerElements={shimmerItem} />
                            <Shimmer shimmerElements={shimmerCategory} />
                            <Shimmer shimmerElements={shimmerItem} />
                            <Shimmer shimmerElements={shimmerItem} />
                            <Shimmer shimmerElements={shimmerCategory} />
                            <Shimmer shimmerElements={shimmerItem} />
                            <Shimmer shimmerElements={shimmerItem} />
                        </Stack>
                    </div>
                </div>
                :
                reportGroups &&
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                        {renderGroup(reportGroups)}
                    </div>
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                        {reportGroups && `Total hours: ${reportGroups.TotalHours()}h`}
                    </div>
                </div>
            }
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <Toggle label="Show task hours"
                        inlineLabel
                        onText="Show"
                        offText="Hide"
                        checked={showTaskHours}
                        onChange={onShowTaskHoursChange} />
                </div>
            </div>

        </div>
    );
};

// export default withAITracking(reactPlugin, TasksReport);
export default TasksReport;