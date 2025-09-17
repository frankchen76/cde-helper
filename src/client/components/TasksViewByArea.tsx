import * as React from "react";
import { useState, useEffect, useContext } from "react";
import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { Dropdown, IDropdownStyles, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { DatePicker } from "@fluentui/react/lib/DatePicker";
import { List } from "@fluentui/react/lib/List";
import { TextField } from "@fluentui/react/lib/TextField";

import moment from "moment";
import { useParams, useHistory } from "react-router-dom";
import { ServiceContext } from "../services/SettingService";
import { Task, TaskCollection } from "../services/Task";
import { TasksRow } from "./TasksRow";
import { IconButton } from "@fluentui/react/lib/Button";
import { cloneDeep } from "lodash"

export interface ITasksViewByAreaProps {
    //outlookItem: OutlookItem;
}

const TasksViewByArea = (props: ITasksViewByAreaProps) => {
    const currentUrl = new URL(window.location.href);
    const hostInfo = currentUrl.searchParams.get("_host_Info");

    const para = useParams();
    const history = useHistory();
    //const settingItemId: string = para["settingItemId"];
    const [selSettingItemId, setSelSettingItemId] = useState<string>(para["settingItemId"]);
    const [selAreaId, setSelAreaId] = useState<number>(+para["areaId"]);
    const [tasks, setTasks] = useState<TaskCollection>();
    const [tasksView, setTasksView] = useState<TaskCollection>();
    const [error, setError] = useState<string>();
    const [loading, setLoading] = useState<boolean>(false);
    const [isDialog, setIsDialog] = useState<boolean>(hostInfo && hostInfo.indexOf("isDialog") != -1);
    const [txtSearch, setTxtSearch] = useState<string>("");
    const serviceContext = useContext(ServiceContext);

    const settingItemsOptions = serviceContext.setting.toIDropdownOption();
    const [areaOptions, setAreaOptions] = useState<IDropdownOption[]>([]);

    //const tempDate = moment().weekday(0).isSame(moment(), "day") ? moment().startOf('day').add(-7, 'days').toDate() : moment().weekday(0).startOf('day').toDate();
    const tempDate = moment().startOf('month').isSame(moment(), "day") ? moment().startOf('day').add(-1, 'months').toDate() : moment().startOf('month').toDate();

    const [startDate, setStartDate] = useState<Date>(tempDate); // always use this monday or previous monday as start date
    console.log("tempDate", tempDate);

    const [endDate, setEndDate] = useState<Date>(new Date());

    // Load tasks
    useEffect(() => {
        const init = async () => {
            try {
                setLoading(true);
                const { taskService, setting, areaService } = serviceContext;
                const selSettingItem = setting.getSettingItemById(selSettingItemId);
                const areas = await areaService.getAreas(selSettingItem);
                setAreaOptions(areas.toIDropdownOption(selSettingItem));
                const selArea = areas.getAreaById(selAreaId);

                await _loadTasks(selSettingItemId, selAreaId, startDate, endDate);
            } catch (error) {
                setError(error);
            }
            setLoading(false);
        };
        init();

    }, []);

    const _loadTasks = async (settingItemId: string, areaId: number, s: Date, e: Date) => {
        const { taskService, setting, areaService } = serviceContext;
        let tasks: TaskCollection = null;
        const selSettingItem = setting.getSettingItemById(settingItemId);
        const areas = await areaService.getAreas(selSettingItem);
        const selArea = areas.getAreaById(areaId);
        const result = await taskService.getTasksByArea(selSettingItem, { areaPath: selArea.WiqlPath, start: s, end: e });
        //tasks = new TaskCollection(result.items);
        //setTasks(result);
        _updateTaskView(result, txtSearch);
    };
    const _updateTaskView = (tasks: TaskCollection, searchTxt: string = null): void => {
        let tView = new TaskCollection(cloneDeep(tasks.items));
        if (searchTxt && searchTxt.length > 0) {
            const filtered = tasks.items.map(t => {
                if ((t.title && t.title.toLowerCase().indexOf(searchTxt.toLowerCase()) != -1) ||
                    (t.issue && t.issue.title && t.issue.title.toLowerCase().indexOf(searchTxt.toLowerCase()) != -1))
                    return t;
            });
            tView = new TaskCollection(filtered);
        }
        setTasks(tasks);
        setTasksView(tView)
    }

    const _onSettingItemChange = async (e, option): Promise<void> => {
        const { taskService, setting } = serviceContext;
        const selSettingItem = setting.getSettingItemById(option.key.toString());
        const areas = await serviceContext.areaService.getAreas(selSettingItem);
        setSelSettingItemId(option.key.toString());
        setAreaOptions(areas.toIDropdownOption(selSettingItem));

    };

    const _onAreaChange = async (e, option): Promise<void> => {
        try {
            setLoading(true);
            setSelAreaId(+option.key);
            await _loadTasks(selSettingItemId, +option.key, startDate, endDate);
        } catch (error) {
            setError(error);
        }
        setLoading(false);
    };

    const _onStartDateChange = async (date: Date | null | undefined): Promise<void> => {
        try {
            setLoading(true);
            setStartDate(date);
            await _loadTasks(selSettingItemId, selAreaId, date, endDate);
        } catch (error) {
            setError(error);
        }
        setLoading(false);
    };

    const _onEndDateChange = async (date: Date | null | undefined): Promise<void> => {
        try {
            setLoading(true);
            setEndDate(date);
            await _loadTasks(selSettingItemId, selAreaId, startDate, date);
        } catch (error) {
            setError(error);
        }
        setLoading(false);
    };

    const onClickTaskHandler = (taskId: number): void => {
        history.push(`/taskitem/${taskId}`);
    }

    const _onIconClickHandler = async (task: Task): Promise<void> => {
        if (task.outlookMessage) {
            await task.outlookMessage.popupForm(isDialog);
        }
    }
    const containerStackTokens: IStackTokens = { childrenGap: 5 };

    const _onRenderCell = (item: Task, index: number, isScrolling: boolean): JSX.Element => {
        return item ? (
            <TasksRow task={item} itemIconClickHandler={_onIconClickHandler} />
        ) : null;
    };
    const _onRenderCell1 = (nestingDepth?: number, item?: Task, itemIndex?: number): React.ReactNode => {

        return item ? (
            <TasksRow task={item} itemIconClickHandler={_onIconClickHandler} />
        ) : null;
    };
    const _onSearchChange = (e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        setTxtSearch(newValue);
        _updateTaskView(tasks, newValue);
    }
    const _onSearchResetClick = (): void => {
        setTxtSearch("");
        _updateTaskView(tasks, null);
    }
    const dropdownStyles: Partial<IDropdownStyles> = {
        dropdownOptionText: { overflow: 'visible', whiteSpace: 'normal' },
        dropdownItem: { height: 'auto' },
    };

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>Tasks</h2>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Dropdown
                        label="Setting"
                        options={settingItemsOptions}
                        onChange={_onSettingItemChange}
                        selectedKey={selSettingItemId}
                        styles={dropdownStyles}
                    />
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Dropdown
                        label="Area"
                        options={areaOptions}
                        onChange={_onAreaChange}
                        selectedKey={selAreaId}
                        styles={dropdownStyles}
                    />
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <DatePicker
                        label="Start Date"
                        placeholder="Select a start date..."
                        ariaLabel="Select a start date"
                        value={startDate}
                        onSelectDate={_onStartDateChange}
                    />
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <DatePicker
                        label="End Date"
                        placeholder="Select a end date..."
                        ariaLabel="Select a end date"
                        value={endDate}
                        onSelectDate={_onEndDateChange}
                    />
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm10 ms-md10 ms-lg10" >
                    <TextField label="Search"
                        placeholder="Search by task or issue title"
                        value={txtSearch}
                        onChange={_onSearchChange} />
                </div>
                <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2" >
                    <IconButton iconProps={{ iconName: 'EraseTool' }}
                        style={{ marginTop: "30px" }}
                        title="Clear the search"
                        ariaLabel="Clear the search"
                        onClick={_onSearchResetClick} />
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <Stack tokens={containerStackTokens}>
                        {loading ? <Spinner title="loading" />
                            :
                            <List items={tasksView && tasksView.items ? tasksView.items : []}
                                onRenderCell={_onRenderCell} />
                        }
                        {error && <div>{error}</div>}
                    </Stack>
                </div>
            </div>
        </div>
    );
};

//export default withAITracking(reactPlugin, TasksView);
export default TasksViewByArea;