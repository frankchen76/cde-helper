import * as React from "react";
import { useState, useEffect, useContext } from "react";
import {
    getFocusStyle,
    getTheme,
    IconButton,
    IIconProps,
    mergeStyleSets,
    Separator,
    Link,
    ITooltipProps,
    TooltipHost,
    IContextualMenuProps,
    CommandButton,
    ITooltipHostStyles,
    Dialog,
    DialogType,
    Dropdown,
    DialogFooter,
    PrimaryButton,
    DefaultButton,
    TextField,
    SpinButton,
    Position,
    CompoundButton,
    Spinner
} from "@fluentui/react";
import * as _ from "lodash";
import { ServiceContext } from "../services/SettingService";
import { OutlookItem } from "../services/OutlookItem";
import { Task } from "../services/Task";
import { TaskStateComponent } from "./TaskStateComponent";
import { Common } from "../services/Common";
import { cloneDeep, set } from "lodash";
import { err } from "../../services/log";
import { useForm } from "react-hook-form";
import { CtrlDropdown, CtrlSpinButton, CtrlTextField, SpinnerButton } from "./ValidationControls";
import { info } from "../services/log";

export interface ITasksViewProps {
    task: Task;
    //outlookItem: OutlookItem;
    itemIconClickHandler: (task: Task) => Promise<void>;
}
type TaskFormInput = {
    stateId: string;
    completedHours: number;
    comments: string;
};

export const TasksRow = (props: ITasksViewProps) => {
    const [task, setTask] = useState(props.task);
    const [isRunning, setIsRunning] = useState<boolean>(false);
    const [prevTask, setPrevTask] = useState<Task>(props.task);
    const [comments, setComments] = useState("");
    const [hideDialog, setHideDialog] = useState(true);
    const serviceContext = useContext(ServiceContext);
    const { control, watch, formState: { errors }, handleSubmit, getValues } = useForm<TaskFormInput>();

    let iconName = "";
    switch (task.outlookMessage.ItemType) {
        case "task":
            iconName = "TaskLogo";
            break;
        case "message":
            iconName = "Mail";
            break;
        case "appointment":
            iconName = "Event";
            break;
    }
    const theme = getTheme();
    const classNames = mergeStyleSets({
        itemCell: [
            getFocusStyle(theme, { inset: -1 }),
            {
                // minHeight: 54,
                // padding: 10,
                // boxSizing: 'border-box',
                // borderBottom: `1px solid ${getTheme().semanticColors.bodyDivider}`,
                // display: 'flex',
                selectors: {
                    '&:hover': { background: theme.palette.neutralLight },
                },
            }
        ],
        selectedRow: {
            backgroundColor: theme.palette.themeLighterAlt
        }
    });

    const iconProps: IIconProps = {
        iconName: iconName,
        style: { fontSize: 15 }
    };
    const [isHighlighted, setIsHighlighted] = useState<boolean>(task && task.outlookMessage && task.outlookMessage.ItemId == serviceContext.selectedOutlookItem?.ItemId);
    useEffect(() => {
        const isH = task && task.outlookMessage && task.outlookMessage.ItemId == serviceContext.selectedOutlookItem?.ItemId;
        setIsHighlighted(isH);
        //console.log("isHighlighted", isH);
    }, [serviceContext.selectedOutlookItem]);

    // set tooltip style
    const tooltipStyle = {
        maxHeight: 200,
        overflow: "auto"
    };
    const tooltipProps: ITooltipProps = {
        onRenderContent: () => (
            <div>
                <div style={tooltipStyle} dangerouslySetInnerHTML={{ __html: `<h2>${task.title}</h2><hr>${task.description}` }}>
                </div>

            </div>
        ),
    };
    const dialogHandler = async () => {
        const { taskService, setting } = serviceContext;

        handleSubmit(
            async (data) => {
                info("dialogHandler:validation successed:", data);
                setIsRunning(true);
                try {
                    let newTask = cloneDeep(task);
                    newTask.state = data.stateId;
                    newTask.completedWork = data.completedHours;

                    await taskService.updateTask(setting.upn, task.settingItem, newTask, data.comments, prevTask);
                    if (newTask.outlookMessage.ItemType == "message") {
                        // update category for message
                        await newTask.outlookMessage.setCategory(newTask.state, false)
                    }

                    setPrevTask(newTask);
                    setTask(newTask);
                    setHideDialog(true);
                    info("dialogHandler:successed:hide dialog");
                } catch (error) {
                    err("dialogHandler:error:", error);
                } finally {
                    setIsRunning(false);
                }
            },
            (error) => {
                err(`dialogHandler:validation:error:`, error);
            }
        )();

    }
    // const dialogStateHandler = (ev, menuItem) => {
    //     let newTask = cloneDeep(task);
    //     newTask.state = menuItem.key;
    //     setTask(newTask);
    // }
    // const statusHandler = (ev, menuItem) => {
    //     const { taskService, setting } = serviceContext;
    //     let newTask = cloneDeep(task);
    //     newTask.state = menuItem.key;
    //     taskService.updateTaskState(task.settingItem, newTask)
    //         .then(() => {
    //             // update category for message
    //             if (newTask.outlookMessage.ItemType == "message") {
    //                 // update category for message
    //                 return newTask.outlookMessage.setCategory(newTask.state, false)
    //             } else
    //                 return Promise.resolve();
    //         })
    //         .catch((error) => {
    //             console.error("Update task state failed", error);
    //         })
    //         .finally(() => {
    //             setTask(newTask);
    //         });

    // };
    const btnContextMenuStyle = {
        height: 20,
        paddingLeft: 0
    };
    const contextMenuItems: IContextualMenuProps = {
        useTargetAsMinWidth: true,
        items: [
            {
                key: "update",
                text: "Update",
                iconProps: { iconName: "EditNote" },
                onClick: (ev, menuItem) => { setHideDialog(false); }
                // subMenuProps: {
                //     items: [
                //         {
                //             key: "ToDo",
                //             text: "To Do",
                //             iconProps: task.state == "To Do" ? { iconName: "StatusCircleCheckmark" } : {},
                //             onClick: (ev, menuItem) => { setHideDialog(false); }
                //         },
                //         {
                //             key: "Doing",
                //             text: "Doing",
                //             iconProps: task.state == "Doing" ? { iconName: "StatusCircleCheckmark" } : {},
                //             onClick: (ev, menuItem) => { setHideDialog(false); }
                //         },
                //         {
                //             key: "Done",
                //             text: "Done",
                //             iconProps: task.state == "Done" ? { iconName: "StatusCircleCheckmark" } : {},
                //             onClick: (ev, menuItem) => { setHideDialog(false); }
                //         }
                //     ]
                // }
            }
        ]
    };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const dialogContentProps = {
        type: DialogType.normal,
        title: 'Update status'
        // subText: 'Your Inbox has changed. No longer does it include favorites, it is a singular destination for your emails.'
    };
    const modelProps = {
        isBlocking: false,
        styles: { main: { maxWidth: 350 } },
    };
    //const cmdStyles:ICommandBarSty
    return task ? (
        <div className={`${isHighlighted ? classNames.selectedRow : classNames.itemCell} ms-Grid`}>
            <div className="ms-Grid-row" style={{ fontSize: "14px" }}>
                <div className="ms-Grid-col ms-sm2">
                    <TooltipHost tooltipProps={tooltipProps}
                        styles={hostStyles}
                        closeDelay={500}
                        calloutProps={{ gapSpace: 0 }}>
                        <IconButton iconProps={iconProps} title={task.outlookMessage ? "Email task" : "Task"} onClick={props.itemIconClickHandler.bind(this, task)} />
                    </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm8 divTaskItem">
                    <Link href={`#/taskitem/${task.settingItem.id}/${task.id}`} >{`${task.title}`}</Link>
                </div>
                <div className="ms-Grid-col ms-sm2">
                    <TaskStateComponent state={task.state} />
                </div>
            </div>
            <div className="ms-Grid-row" style={{ fontSize: "12px" }}>
                <div className="ms-Grid-col ms-sm2">
                    <span style={{ fontWeight: "bold" }}>Issue:</span>
                </div>
                <div className="ms-Grid-col ms-sm8">
                    <span style={{ fontWeight: "bold" }}>{`${task.issue && task.issue.title}`}</span>
                </div>
                <div className="ms-Grid-col ms-sm2" style={{ paddingLeft: 0 }}>
                    {/* <DefaultButton text="" menuProps={contextMenuItems} /> */}
                    <CommandButton style={btnContextMenuStyle} text="" menuProps={contextMenuItems} />
                    {/* <CommandBar
                        items={[]}
                        styles={{ root: { padding: 0, height: 16 } }}
                        overflowItems={_overflowItems}
                    /> */}
                    <Dialog
                        hidden={hideDialog}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: task.title
                            // subText: 'Your Inbox has changed. No longer does it include favorites, it is a singular destination for your emails.'
                        }}
                        modalProps={modelProps}
                        onDismiss={() => setHideDialog(true)}
                    >
                        {task && <CtrlDropdown
                            label="State:"
                            name="stateId"
                            rules={{ required: "State is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            options={Common.getStateOptions()}
                            selectedKey={task.state}
                        />}
                        {task && <CtrlSpinButton
                            label={'Completed Hour:'}
                            name="completedHours"
                            rules={{ required: "Completed hour is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={task.completedWork.toString()}
                            min={0}
                            max={16}
                            step={0.25}
                        />}

                        {task && <CtrlTextField
                            label="Comments:"
                            name="comments"
                            rules={undefined}
                            control={control}
                            errors={errors}
                            required={false}
                            multiline={true}
                            defaultValue={comments} />}

                        <DialogFooter>
                            <SpinnerButton onClick={dialogHandler} text="Save" isRunning={isRunning} />
                            <DefaultButton onClick={() => setHideDialog(true)} text="Cancel" />
                        </DialogFooter>
                    </Dialog>
                </div>
            </div>
            <div className="ms-Grid-row" style={{ fontSize: "11px" }}>
                <div className="ms-Grid-col ms-sm2">
                    <span style={{ fontWeight: "bold" }}>ID:</span>
                </div>
                <div className="ms-Grid-col ms-sm4">
                    <span style={{ fontWeight: "bold" }}>Created: </span>
                </div>
                <div className="ms-Grid-col ms-sm4">
                    <span style={{ fontWeight: "bold" }}>Due: </span>
                </div>
                <div className="ms-Grid-col ms-sm2">
                    <span style={{ fontWeight: "bold" }}>Hrs: </span>
                </div>
            </div>
            <div className="ms-Grid-row" style={{ fontSize: "11px" }}>
                <div className="ms-Grid-col ms-sm2">
                    <Link href={task.urlHtml} >{`${task.id}`}</Link>
                </div>
                <div className="ms-Grid-col ms-sm4">
                    {Common.dateToDurationString(task.createdDate, task.id)}
                </div>
                <div className="ms-Grid-col ms-sm4">
                    {Common.dateToDurationString(task.dueDate, task.id)}
                </div>
                <div className="ms-Grid-col ms-sm2">
                    {task.completedWork}
                </div>
            </div>
            <Separator styles={{ root: { padding: 0, height: "10px" } }} />
        </div>
    ) : null;
}