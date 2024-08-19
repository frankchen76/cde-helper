import * as React from "react";
import { useState, useEffect, useContext } from "react";
import {
    getFocusStyle,
    getTheme,
    SelectionMode,
    Stack,
    IStackTokens,
    IGroup,
    IconButton,
    IIconProps,
    GroupedList,
    mergeStyleSets,
    Separator,
    Link,
    Spinner,
    HoverCard,
    IExpandingCardProps,
    ExpandingCardMode,
    ITooltipProps,
    TooltipHost,
    CommandBar,
    ICommandBarItemProps,
    DefaultButton,
    IContextualMenuProps,
    CommandButton
} from "@fluentui/react";
import { RouteComponentProps } from "react-router-dom";
import { ServiceContext } from "../services/SettingService";
import { OutlookItem } from "../services/OutlookItem";
import { Task, TaskCollection } from "../services/Task";
// import { withAITracking } from '@microsoft/applicationinsights-react-js';
// import { reactPlugin, appInsights } from '../services/AppInsights';
import { TaskStateComponent } from "./TaskStateComponent";
import { Common } from "../services/Common";
import { cloneDeep } from "lodash";
import { TasksRow } from "./TasksRow";

export interface ITasksViewProps {
    routeProps: RouteComponentProps;
    //currentMailItem: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead;
    //outlookItem: OutlookItem;
}

const TasksView = (props: ITasksViewProps) => {
    const currentUrl = new URL(window.location.href);
    const hostInfo = currentUrl.searchParams.get("_host_Info");

    const [tasks, setTasks] = useState<TaskCollection>();
    const [groups, setGroups] = useState<IGroup[]>();
    const [error, setError] = useState<string>();
    const [loading, setLoading] = useState<boolean>(false);
    const [isDialog, setIsDialog] = useState<boolean>(hostInfo && hostInfo.indexOf("isDialog") != -1);
    const serviceContext = useContext(ServiceContext);

    // Load tasks
    useEffect(() => {
        const loadTasks = async () => {
            try {
                setLoading(true);
                const { taskService, setting } = serviceContext;
                let tasks: TaskCollection = null;
                for (const settingItem of setting.items) {
                    const result = await taskService.getTasksByQueryId(settingItem);
                    if (result && result.items && result.items.length > 0) {
                        //sort the return tasks
                        result.sort();

                        if (tasks == null) {
                            tasks = new TaskCollection(result.items);
                        } else {
                            tasks.appendTasks(result);
                        }
                        //console.log(`settingItem: ${settingItem.id}; task count: ${tasks.items.length}`);
                    }
                }
                if (tasks && tasks.items && tasks.items.length > 0) {
                    setTasks(tasks);
                } else {
                    setError("No tasks was found.");
                }
            } catch (error) {
                setError(error);
            }
            setLoading(false);
        };
        loadTasks();
    }, []);

    // Get Outlook item
    // useEffect(() => {
    //     const loadOutlookItem = async () => {
    //         if (props.currentMailItem) {
    //             const newItem = await OutlookItem.createInstance(props.currentMailItem);
    //             setOutlookItem(newItem);
    //         }
    //     };
    //     loadOutlookItem();
    // }, [props.currentMailItem]);

    //Load groups
    useEffect(() => {
        if (tasks) {
            const groups = tasks.createGroupsFromTasks();
            setGroups(groups);
        }
    }, [tasks]);

    // refresh Outlook item
    // useEffect(() => {
    //     setSelectedOutlookItem(props.outlookItem);
    //     console.log("Outlook item changed from parent.");
    // }, [props.outlookItem]);

    const onClickTaskHandler = (taskId: number): void => {
        props["routeProps"]["history"].push(`/taskitem/${taskId}`);
    }

    const _onIconClickHandler = async (task: Task): Promise<void> => {
        if (task.outlookMessage) {
            await task.outlookMessage.popupForm(isDialog);
        }
    }
    const containerStackTokens: IStackTokens = { childrenGap: 5 };

    const _onRenderCell = (nestingDepth?: number, item?: Task, itemIndex?: number): React.ReactNode => {
        return item ? (
            <TasksRow task={item} itemIconClickHandler={_onIconClickHandler} />
        ) : null;
    };

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>Tasks</h2>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    {loading ? <Spinner title="loading" />
                        :
                        // <ShimmeredDetailsList
                        //     className={"divTaskView"} //fixed the scollbar issue
                        //     indentWidth={0}
                        //     cellStyleProps={{ cellLeftPadding: 5, cellRightPadding: 0, cellExtraRightPadding: 0 }}
                        //     compact={true}
                        //     setKey="items"
                        //     items={tasks && tasks.items ? tasks.items : []}
                        //     columns={columns}
                        //     groups={groups}
                        //     onRenderItemColumn={_renderItemColumn.bind(this)}
                        //     selectionMode={SelectionMode.none}
                        //     enableShimmer={!tasks}
                        //     onRenderRow={onRenderRow.bind(this)}
                        //     ariaLabelForShimmer="Content is being fetched"
                        //     ariaLabelForGrid="Item details"
                        // />
                        <GroupedList items={tasks && tasks.items ? tasks.items : []}
                            onRenderCell={_onRenderCell}
                            // groupProps={{ onRenderHeader: _onRenderHeader }}
                            selectionMode={SelectionMode.none}
                            compact={true}
                            groups={groups} />
                    }
                    {error && <div>{error}</div>}
                </div>
            </div>
        </div>
    );
};

//export default withAITracking(reactPlugin, TasksView);
export default TasksView;