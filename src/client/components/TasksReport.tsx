import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
// import { withAITracking } from "@microsoft/applicationinsights-react-js";
// import { reactPlugin } from "../services/AppInsights";
import { useContext, useEffect, useState } from "react";
import { Common, ExecutingResult } from "../services/Common";
import { ServiceContext } from "../services/SettingService";
import { ReportItem, ReportItemCollection, ReportItemGroup, ReportItemGroupCollection } from "../services/ReportItem";

import { IconButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Shimmer, ShimmerElementType } from "@fluentui/react/lib/Shimmer";
import { Stack } from "@fluentui/react/lib/Stack";
import { Toggle } from "@fluentui/react/lib/Toggle";
import { DatePicker } from "@fluentui/react/lib/DatePicker";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import moment from "moment";
import { cloneDeep, find, findIndex } from "lodash";

export interface ITaskReportItemProps {
    reportItem: ReportItem;
    showTaskHours: boolean;
    onReportItemChanged: (item: ReportItem) => void;
}
const TaskReportItem = (props: ITaskReportItemProps) => {
    const [reportItem, setReportItem] = useState<ReportItem>(props.reportItem);
    const onRecordedChanged = (ev, checked: boolean) => {
        //console.log(areaSetting);
        const newReportItem = cloneDeep(props.reportItem);
        newReportItem.Recorded = checked;
        setReportItem(newReportItem);
        props.onReportItemChanged(newReportItem);
    };

    return (<li key={`li1-${reportItem.Id.toString()}`} id={reportItem.Id.toString()}>
        <Checkbox key={`chkEnabled_${reportItem.Id.toString()}`}
            label={`${reportItem.Title} ${props.showTaskHours ? "(" + reportItem.TodayHours + ")" : ""}`}
            checked={reportItem.Recorded}
            onChange={onRecordedChanged} />

    </li>);
};

export interface ITasksReportProps {
    routeProps: RouteComponentProps;
}
const TasksReport = (props: ITasksReportProps) => {
    const [executingResult, setExecutingResult] = useState<ExecutingResult>(ExecutingResult.createInstance());
    const [reportItems, setReportItems] = useState<ReportItemCollection>();
    const [reportGroups, setReportGroups] = useState<ReportItemGroupCollection>();
    const [showTaskHours, setShowTaskHours] = useState<boolean>(true);
    const [reportDate, setReportDate] = useState<Date>(new Date());
    const serviceContext = useContext(ServiceContext);

    useEffect(() => {
        loadReports();
    }, []);

    const loadReports = async (reportDate?: string) => {
        try {
            const { reportService, setting } = serviceContext;
            let allReportItems: ReportItemCollection = null;
            let dbResult = "";
            setExecutingResult(ExecutingResult.start());

            if (reportDate == undefined) {
                allReportItems = await reportService.getReportItems(setting);

                // Log report items to DB
                if (allReportItems && allReportItems.Items && allReportItems.Items.length > 0) {
                    dbResult = await reportService.logReportItemsToDb(allReportItems, null);
                }
            } else {
                allReportItems = await reportService.loadReportItemsFromDb(reportDate);
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

    const onReportItemChangedHandler = (item: ReportItem): void => {
        // const newReportItems = cloneDeep(reportItems);
        // const foundIndex = findIndex(newReportItems.Items, i => i.Id === item.Id);
        // if (foundIndex > -1) {
        //     newReportItems.Items[foundIndex].Recorded = item.Recorded;
        // }
        let newReportItems = new ReportItemCollection(reportItems.Items.map(i => {
            if (i.Id === item.Id) {
                i.Recorded = item.Recorded;
            }
            return i;
        }));
        setReportItems(newReportItems);
    };

    const renderTask = (group: ReportItemGroup): any => {
        return (
            <ul key={`ul-${group.GroupName}`} id={`ul-${group.GroupName}`} style={{ paddingLeft: "20px" }}>
                {group.ReportItems.map(reportItem => {
                    return <TaskReportItem key={`tri-${reportItem.Id}`} reportItem={reportItem} showTaskHours={showTaskHours} onReportItemChanged={onReportItemChangedHandler} />;
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
                    return (<li key={`li2-${header}`} id={`li-${header}`}>
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
    const onSelectDate = async (date: Date | null | undefined): Promise<void> => {
        setReportDate(date);
        const selDate = date ? moment(date) : moment();
        if (selDate.isSame(new Date(), "day")) {
            await loadReports();
        } else {
            //await loadReports(date ? date.toISOString().substring(0, 10) : undefined);
            await loadReports(selDate.format("YYYY-MM-DD"));
        }
    }
    const onUpdateClick = async (): Promise<void> => {
        try {
            const { reportService, setting } = serviceContext;
            let dbResult = "";
            setExecutingResult(ExecutingResult.start());

            // Update report items to DB
            if (reportItems && reportItems.Items && reportItems.Items.length > 0) {
                dbResult = await reportService.logReportItemsToDb(reportItems, moment(reportDate).format("YYYY-MM-DD"));
            }

            if (dbResult === "") {
                setExecutingResult(ExecutingResult.complete(false));
            } else {
                setExecutingResult(ExecutingResult.complete(true, dbResult, true));
            }

        } catch (error) {
            setExecutingResult(ExecutingResult.complete(true, error, true));
        }

    }

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>Tasks Report</h2>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <DatePicker
                        isRequired={true}
                        today={new Date()}
                        label="Report Date:"
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        maxDate={new Date()}
                        value={reportDate}
                        onSelectDate={onSelectDate}
                    />
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
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <PrimaryButton text="Update" onClick={onUpdateClick} />
                </div>
            </div>

        </div>
    );
};

// export default withAITracking(reactPlugin, TasksReport);
export default TasksReport;