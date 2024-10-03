import {
    DefaultButton,
    IDropdownOption,
    MessageBar,
    MessageBarType,
    PrimaryButton,
    Shimmer,
    Stack,
} from "@fluentui/react";
import * as React from "react";
import * as _ from "lodash";
import { useParams, useHistory } from "react-router-dom";
//import 'react-quill/dist/quill.snow.css';
import { Issue, IssueCollection } from "../services/Issue";
import { Task } from "../services/Task";
import { ISettingItem, ServiceContext } from "../services/SettingService";
import { TagCollection } from "../services/Tag";
import { OutlookItem, OutlookItemJSON } from "../services/OutlookItem";
import { Area, AreaCollection } from "../services/Area";
// import { withAITracking } from '@microsoft/applicationinsights-react-js';
// import { reactPlugin, appInsights } from '../services/AppInsights';
import { useContext, useEffect, useState } from "react";
import { ShimmerButtons, ShimmerCtrl, ShimmerQuill, TaskItemShimmer } from "./TaskItemShimmer";
import { Common, ExecutingResult, StateEnum, TaskFormModeEnum } from "../services/Common";
import { useForm, SubmitHandler, Controller } from "react-hook-form";
import { CtrlDatePicker, CtrlDropdown, CtrlQuill, CtrlSpinButton, CtrlTaskTagPicker, CtrlTextField, CtrlTextField2, CtrlWatchDropdown, ICtrlWatchDropdownOptions, SpinnerButton } from "./ValidationControls";
import moment from "moment";

type TaskFormInput = {
    title: string;
    settingItemId: string;
    areaId: number;
    issueId: number;
    stateId: string;
    completedHours: number;
    tags: string;
    description: string;
    dueDate: Date;
};

export interface ITaskItemViewProps {
    // settingService: SettingService;
    // routeProps: RouteComponentProps;
    //currentMailItem: Office.Item & Office.ItemCompose & Office.ItemRead & Office.Message & Office.MessageCompose & Office.MessageRead & Office.Appointment & Office.AppointmentCompose & Office.AppointmentRead;
    outlookItem: OutlookItem;
}

const TaskItemView = (props: ITaskItemViewProps) => {

    const [task, setTask] = useState<Task>();
    const [prevTask, setPrevTask] = useState<Task>();
    const [executingResult, setExecutingResult] = useState<ExecutingResult>(ExecutingResult.createInstance());
    const serviceContext = useContext(ServiceContext);

    const currentUrl = new URL(window.location.href);
    const hostInfo = currentUrl.searchParams.get("_host_Info");
    const isDialog = hostInfo && hostInfo.indexOf("isDialog") != -1;

    const para = useParams();
    const history = useHistory();
    const id: number = +para["id"];
    const settingItemId: string = para["settingItemId"];

    const [settingItem, setSettingItem] = useState<ISettingItem>(serviceContext.setting.getSettingItemById(settingItemId));
    const [formMode, setFormMode] = useState<TaskFormModeEnum>(Common.getTaskFormModeFromId(id));
    // const [settingItem, setSettingItem] = useState<ISettingItem>();
    // const [formMode, setFormMode] = useState<TaskFormModeEnum>();

    const stateOptions = Common.getStateOptions();

    const { control, watch, formState: { errors }, handleSubmit, getValues } = useForm<TaskFormInput>();

    // const [areas, setAreas] = useState<AreaCollection>();
    // const [selectedArea, setSelectedArea] = useState<Area>();
    const [outlookItem, setOutlookItem] = useState<OutlookItem>();

    // Load tags
    const [tags, setTags] = useState<TagCollection>();

    // Load task
    useEffect(() => {
        const loadTask = async () => {
            try {
                setExecutingResult(ExecutingResult.start());
                const { taskService, areaService, issueService, tagService } = serviceContext;
                let tags = await tagService.getTags(settingItem)

                let task: Task = null;
                switch (formMode) {
                    case TaskFormModeEnum.CreateEmailTask:
                        task = new Task();
                        task.id = 0;
                        task.title = props.outlookItem.Subject;
                        task.description = "";
                        task.state = StateEnum.ToDo;;
                        task.completedWork = 0;
                        task.tags = undefined;

                        task.areaId = undefined;
                        task.issue = undefined;
                        task.outlookMessage = props.outlookItem.toOutlookItemJSON();
                        if (props.outlookItem.ItemType == "appointment") {
                            task.completedWork = (Math.round(props.outlookItem.Duration.asHours() * 4) / 4);
                            task.description = props.outlookItem.Subject;
                            task.tags = "Meeting";
                            task.state = StateEnum.Done;
                        }
                        //task.dueDate = moment(new Date()).startOf('day').add(3, 'days').toDate()
                        task.dueDate = null;
                        break;
                    case TaskFormModeEnum.CreateTask:
                        //init regular task
                        task = new Task();
                        task.id = 0;
                        task.title = "";
                        task.description = "";
                        task.state = StateEnum.ToDo;
                        task.completedWork = 0;
                        task.tags = undefined;
                        task.areaId = undefined;
                        task.area = undefined;
                        task.issue = undefined;
                        task.outlookMessage = OutlookItemJSON.createTaskInstance();
                        //task.dueDate = moment(new Date()).startOf('day').add(3, 'days').toDate()
                        task.dueDate = null;
                        break;
                    case TaskFormModeEnum.UpdateTask:
                        task = await taskService.getTaskById(settingItem, id);
                        break;
                }
                // copy the exist one as original
                setOutlookItem(props.outlookItem);
                setPrevTask(_.cloneDeep(task));
                setTags(tags);
                setTask(task);
                setExecutingResult(ExecutingResult.complete());

            } catch (error) {
                console.log(error);
                setTask(null);
                setExecutingResult(ExecutingResult.complete(true, Common.getErrorMessage(error), true));
                //setMessage(error);
            }

        };

        loadTask();
    }, []);

    const onAddHandler = async () => {
        const { taskService, setting } = serviceContext;
        // setIsExcuting(true);
        handleSubmit(
            async (data) => {
                console.log("validation successed:", data);
                try {
                    setExecutingResult(ExecutingResult.start());
                    const newSettingItem = setting.getSettingItemById(data.settingItemId)
                    let taskValue = _.cloneDeep(task);
                    taskValue.title = data.title;
                    taskValue.settingItem = newSettingItem;
                    //taskValue.area = data.areaId.Path;
                    taskValue.areaId = data.areaId;
                    // TODO
                    const selIssue = await serviceContext.issueService.getIssueById(newSettingItem, data.issueId);
                    taskValue.issue = selIssue;

                    taskValue.state = data.stateId;
                    taskValue.tags = data.tags;
                    taskValue.completedWork = data.completedHours;
                    taskValue.description = data.description;
                    taskValue.dueDate = data.dueDate;

                    // Update Azure DevOps
                    let newTask = await taskService.updateTask(setting.upn, newSettingItem, taskValue, null, prevTask);
                    newTask.outlookMessage = task.outlookMessage;

                    let categoryResult = "";
                    if (newTask.outlookMessage.ItemType == "message") {
                        // update category for message
                        categoryResult = await newTask.outlookMessage.setCategory(taskValue.state, isDialog);
                    }

                    setPrevTask(_.cloneDeep(newTask));
                    setTask(newTask);
                    //setIsUpdating(false);
                    setFormMode(TaskFormModeEnum.UpdateTask);
                    //setMessage("task was updated");
                    setExecutingResult(ExecutingResult.complete(true, `Task was updated ${categoryResult}`));
                } catch (error) {
                    console.log("update", error);
                    //setMessage(error);
                    setExecutingResult(ExecutingResult.complete(true, Common.getErrorMessage(error), true));
                }
            },
            (err) => {
                console.log(`validation error:`, err);
                //setValidationError(err);
                // setIsExcuting(false);
            }
        )();
    };
    const onCancelHandler = () => {
        // props.routeProps.history.push("/");
        history.push("/");
    }
    const onViewEmailHandler = async (): Promise<void> => {
        if (task.outlookMessage) {
            await task.outlookMessage.popupForm(isDialog);
        }
    }
    const onMessageBarDismiss = () => {
        setExecutingResult(result => ({ ...result, displayMessage: false }));
    };
    const loadSetting = async (settingItemId: string): Promise<ISettingItem> => {
        return new Promise((resolve, reject) => {
            let ret = null;
            if (settingItemId) {
                ret = serviceContext.setting.getSettingItemById(settingItemId);
            }
            resolve(ret);
        });

    };
    const loadSettings = async (...param: string[]): Promise<ICtrlWatchDropdownOptions> => {
        return new Promise((resolve, reject) => {
            const options = serviceContext.setting.toIDropdownOption();
            resolve({
                options,
                defaultKey: options && options.length > 0 ? options[0].key : undefined
            });
        });

    };
    const loadAreas = async (settingItemId: string): Promise<ICtrlWatchDropdownOptions> => {
        let ret: ICtrlWatchDropdownOptions = null;
        if (settingItemId) {
            const selSettingItem = serviceContext.setting.getSettingItemById(settingItemId);
            const areas = await serviceContext.areaService.getAreas(selSettingItem);
            const options = areas.toIDropdownOption(selSettingItem);
            const defaultArea = serviceContext.areaService.getAreaByOutlookItem(selSettingItem, areas, outlookItem);
            ret = {
                options,
                defaultKey: defaultArea.Id
            };
        }
        return ret;
    };
    const loadIssuess = async (settingItemId: string, areaId: string): Promise<ICtrlWatchDropdownOptions> => {
        let ret: ICtrlWatchDropdownOptions = null;
        if (settingItemId != undefined && areaId != undefined) {
            const selSettingItem = serviceContext.setting.getSettingItemById(settingItemId);
            //const allArea = await serviceContext.areaService.getAreas(selSettingItem);
            const selArea = await serviceContext.areaService.getAreaById(selSettingItem, +areaId);
            const allIssues = await serviceContext.issueService.getIssuesByArea(selSettingItem, selArea);
            const options = allIssues.toIDropdownOption();
            ret = {
                options,
                defaultKey: options && options.length > 0 ? options[0].key : undefined
            };
        }
        return ret;

    };

    const printTitle = (fm: TaskFormModeEnum): string => {
        let ret = "";
        switch (fm) {
            case TaskFormModeEnum.CreateEmailTask:
                ret = "Create Email Task";
                break;
            case TaskFormModeEnum.CreateTask:
                ret = "Create Task";
                break;
            case TaskFormModeEnum.UpdateTask:
                ret = "Update Task";
                break;

            default:
                break;
        }
        return ret;
    };

    // return !task ? (<TaskItemShimmer />) : (
    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>{printTitle(formMode)}</h2>
                </div>
            </div>
            {executingResult.displayMessage && <div className="ms-Grid-row"><div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                <MessageBar messageBarType={executingResult.isError ? MessageBarType.error : MessageBarType.success}
                    onDismiss={onMessageBarDismiss}
                    isMultiline={false}>{executingResult.message}</MessageBar>
            </div></div>}
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlTextField
                            label="Title:"
                            name="title"
                            rules={{ required: "Title is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            multiline={false}
                            defaultValue={task.title} />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlWatchDropdown
                            label="Setting:"
                            name="settingItemId"
                            rules={{ required: "Setting is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={settingItem.id}
                            disabled={formMode == TaskFormModeEnum.UpdateTask}
                            onGetOptions={loadSettings}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlWatchDropdown
                            label="Area:"
                            name="areaId"
                            rules={{ required: "Area is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={task ? task.areaId : undefined}
                            onGetOptions={loadAreas}
                            watchedName1="settingItemId"
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlWatchDropdown
                            label="Issue:"
                            name="issueId"
                            rules={{ required: "Issue is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={task && task.issue ? task.issue.id : undefined}
                            onGetOptions={loadIssuess}
                            watchedName1="settingItemId"
                            watchedName2="areaId"
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlDropdown
                            label="State:"
                            name="stateId"
                            rules={{ required: "State is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            options={stateOptions}
                            selectedKey={task.state}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
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
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlTaskTagPicker
                            label="Choose tags:"
                            name="tags"
                            rules={{ required: "Tags is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            watchedNameSettingItem="settingItemId"
                            onGetSettingItem={loadSetting}
                            defaultValue={task.tags}
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={task != null}>
                        {task && <CtrlDatePicker
                            label="Due date:"
                            name="dueDate"
                            rules={undefined}
                            control={control}
                            errors={errors}
                            required={false}
                            defaultValue={task.dueDate}
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerQuill()} isDataLoaded={task != null}>
                        {task && <CtrlQuill
                            label="Description:"
                            name="description"
                            rules={{ required: "Description is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            html={task.description}
                            placeholder="Enter description or click 'mail' button to copy selected email body."
                            outlookItem={outlookItem}
                            watchedTitleName="title"
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerButtons()} isDataLoaded={task != null}>
                        <Stack horizontal horizontalAlign="space-around" style={{ marginTop: "10px" }}>
                            <SpinnerButton isRunning={executingResult.isRunning} text={formMode == TaskFormModeEnum.UpdateTask ? "Update" : "Add"} onClick={onAddHandler} />
                            {task && task.outlookMessage && formMode == TaskFormModeEnum.UpdateTask &&
                                <DefaultButton disabled={executingResult.isRunning} text={task.outlookMessage.ItemType == "message" ? "View Email" : "View Event"} onClick={onViewEmailHandler} />
                            }
                            <DefaultButton disabled={executingResult.isRunning} text="Cancel" onClick={onCancelHandler} />
                        </Stack>
                    </Shimmer>
                </div>
            </div>
        </div>
    );
};
//export default withAITracking(reactPlugin, TaskItemView);
export default TaskItemView;