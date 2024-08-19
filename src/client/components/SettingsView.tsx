import { cloneDeep } from "lodash";
import { DefaultButton, Dropdown, IDropdownOption, IStackTokens, Label, MessageBar, MessageBarType, PrimaryButton, SelectableOptionMenuItemType, Shimmer, Stack, StackItem, TextField } from "@fluentui/react";
import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import { QueryService } from "../services/QueryService";
import { ISettingAreaItem, ISettingItem, ServiceContext } from "../services/SettingService";
import { useEffect, useState } from "react";
import { Common, ExecutingResult } from "../services/Common";
import { QueryCollection } from "../services/Query";
// import { withAITracking } from "@microsoft/applicationinsights-react-js";
// import { reactPlugin } from "../services/AppInsights";
import { useForm } from "react-hook-form";
import { CtrlAreaSetting, CtrlDropdown, CtrlTextField2, CtrlToggle } from "./ValidationControls";
import { ShimmerCtrl } from "./TaskItemShimmer";
import * as _ from "lodash";
import { AreaCollection } from "../services/Area";
import { err, info } from "../../services/log";

export interface ISettingsViewProps {
    routeProps: RouteComponentProps;
}

const SettingsView = (props: ISettingsViewProps) => {
    const serviceContext = React.useContext(ServiceContext);
    const { setting, queryService } = serviceContext;
    const [upn, setUpn] = useState<string>(setting.upn);
    const [apiKey, setApiKey] = useState<string>(setting.apiKey);
    const [executingResult, setExecutingResult] = useState<ExecutingResult>(ExecutingResult.createInstance());

    const [currentSettingItemId, setCurrentSettingItemId] = useState<string>(setting.defaultSettingId);

    const onSettingItemsChanged = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        setCurrentSettingItemId(item.key.toString());
    };
    const onUpdateHandler = async (): Promise<void> => {
    };
    const onNewHandler = () => {
    };
    const onRemoveHandler = () => {
    };
    const onCancelHandler = () => {
        props.routeProps.history.push("/");
    };
    const onCleanRoamSettingHandler = async (): Promise<void> => {
        try {
            //const result = await serviceContext.settingService.removeSettingRoam();
            //this.setState({ isUpdating: true });
            //const result = await this.props.settingService.removeSetting();
            //this.setState({ isUpdating: false, message: `Setting was removed ${result ? "successfully" : "failed"}` });
            //await serviceContext.settingService.rem
        } catch (error) {
            //this.setState({ isUpdating: false, message: error });
            console.log(error);
        }
    }
    const onUpnChange = (event, newVal) => {
        setUpn(newVal);
    };
    const onApiKeyChange = async (event, newVal) => {
        setApiKey(newVal);
    };
    const onUpnApplyHandler = async () => {
        setExecutingResult(ExecutingResult.start());
        try {

            let newSetting = _.cloneDeep(setting);
            newSetting.upn = upn;
            // Save the setting to localStorage
            await newSetting.saveSetting()
            serviceContext.onSettingUpdate(newSetting);
            setExecutingResult(ExecutingResult.complete(true, "UPN was applied"));
        } catch (error) {
            console.log("update upn", error);
            //setMessage(error);
            setExecutingResult(ExecutingResult.complete(true, Common.getErrorMessage(error), true));
        }
    };
    const onApiKeyApplyHandler = async () => {
        setExecutingResult(ExecutingResult.start());
        try {

            let newSetting = _.cloneDeep(setting);
            newSetting.apiKey = apiKey;
            // Save the setting to localStorage
            await newSetting.saveSetting()
            serviceContext.onSettingUpdate(newSetting);
            setExecutingResult(ExecutingResult.complete(true, "ApiKey was applied"));
        } catch (error) {
            console.log("update upn", error);
            //setMessage(error);
            setExecutingResult(ExecutingResult.complete(true, Common.getErrorMessage(error), true));
        }
    };
    const onMessageBarDismiss = () => {
        setExecutingResult(result => ({ ...result, displayMessage: false }));
    };

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>Settings</h2>
                </div>
            </div>
            {executingResult.displayMessage && <div className="ms-Grid-row"><div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                <MessageBar messageBarType={executingResult.isError ? MessageBarType.error : MessageBarType.success}
                    onDismiss={onMessageBarDismiss}
                    isMultiline={false}>{executingResult.message}</MessageBar>
            </div></div>}
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Stack horizontal verticalAlign="end" horizontalAlign="space-between" tokens={Common.CONTAINER_STACK_TOKENS}>
                        <Stack.Item grow={2}>
                            <TextField label="UPN:" value={upn} onChange={onUpnChange} />
                        </Stack.Item>
                        <Stack.Item grow={1}>
                            <PrimaryButton disabled={false} text="Apply" onClick={onUpnApplyHandler} />
                        </Stack.Item>
                    </Stack>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Stack horizontal verticalAlign="end" horizontalAlign="space-between" tokens={Common.CONTAINER_STACK_TOKENS}>
                        <Stack.Item grow={2}>
                            <TextField label="ApiKey:" value={apiKey} onChange={onApiKeyChange} />
                        </Stack.Item>
                        <Stack.Item grow={1}>
                            <PrimaryButton disabled={false} text="Apply" onClick={onApiKeyApplyHandler} />
                        </Stack.Item>
                    </Stack>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Stack horizontal verticalAlign="end" horizontalAlign="space-between" tokens={Common.CONTAINER_STACK_TOKENS}>
                        <Stack.Item grow={2}>
                            <Dropdown label="Setting Name" options={setting.toIDropdownOption()} defaultSelectedKey={currentSettingItemId} onChange={onSettingItemsChanged} />
                        </Stack.Item>
                        <Stack.Item grow={1}>
                            <PrimaryButton disabled={false} text="New" onClick={onNewHandler} />
                        </Stack.Item>
                        <Stack.Item grow={1}>
                            <PrimaryButton disabled={false} text="Clean" onClick={onCleanRoamSettingHandler} />
                        </Stack.Item>
                    </Stack>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <SettingItemView settingItemId={currentSettingItemId}
                        onRemove={onRemoveHandler}
                        onUpdate={onUpdateHandler} />
                </div>
            </div>
        </div>
    );
};

interface ISettingItemViewProps {
    settingItemId: string;
    onUpdate: (newSettingItem: ISettingItem) => void;
    onRemove: (settingItemId: string) => void;
}
type SettingItemFormInput = {
    name: string;
    isDefault: boolean;
    tasksQueryId: string;
    issuesQueryId: string;
    reportQueryId: string;
    baseUrl: string;
    areas: ISettingAreaItem[];
};
const SettingItemView = (props: ISettingItemViewProps) => {
    const serviceContext = React.useContext(ServiceContext);
    const { setting, queryService } = serviceContext;
    const [executingResult, setExecutingResult] = useState<ExecutingResult>(ExecutingResult.createInstance());
    const [currentQueries, setCurrentQueries] = useState<QueryCollection>();
    const [currentSettingItem, setcurrentSettingItem] = useState<ISettingItem>();

    const { control, watch, formState: { errors }, handleSubmit, getValues } = useForm<SettingItemFormInput>();

    useEffect(() => {
        const loadQuery = async () => {
            setExecutingResult(ExecutingResult.start());
            try {
                const settingItem = setting.getSettingItemById(props.settingItemId);
                const queries = await queryService.getQueries(settingItem.baseUrl);

                setcurrentSettingItem(settingItem);
                setCurrentQueries(queries);
                setExecutingResult(ExecutingResult.complete());
            } catch (error) {
                setExecutingResult(ExecutingResult.complete(true, error, true));
            }
        };

        loadQuery();
    }, [props.settingItemId]);
    const onBaseUrlValidation = async (val: string): Promise<boolean | string> => {
        let ret: boolean | string = false;
        try {
            const queries = await queryService.getQueries(val);
            ret = true;
        } catch (error) {
            ret = "URL isn't correct";
        }
        return ret;
    };
    const onUpdateHandler = async () => {
        handleSubmit(
            async (data) => {
                console.log("Setting Item form validation successed:", data);
                setExecutingResult(ExecutingResult.start());
                try {
                    let newSetting = _.cloneDeep(setting);
                    if (props.settingItemId == "0") {
                        // when new SettingItem is added
                        newSetting.items.push({
                            id: data.name,
                            name: data.name,
                            baseUrl: data.baseUrl,
                            tasksQueryId: data.tasksQueryId,
                            issuesQueryId: data.issuesQueryId,
                            reportQueryId: data.reportQueryId,
                            areas: data.areas
                        });
                    } else {
                        // when update the existing SettingItem
                        let existSetting = newSetting.getSettingItemById(props.settingItemId);
                        existSetting.id = data.name;
                        existSetting.name = data.name;
                        existSetting.baseUrl = data.baseUrl;
                        existSetting.tasksQueryId = data.tasksQueryId;
                        existSetting.issuesQueryId = data.issuesQueryId;
                        existSetting.reportQueryId = data.reportQueryId;
                        existSetting.areas = data.areas;
                        if (data.isDefault) {
                            newSetting.defaultSettingId = data.name;
                        }
                    }
                    // Save the setting to localStorage
                    await newSetting.saveSetting();
                    info("SettingsView:onUpdateHandler:update successed");
                    serviceContext.onSettingUpdate(newSetting);
                    setExecutingResult(ExecutingResult.complete(true, "Setting was updated"));
                } catch (error) {
                    err("SettingsView:onUpdateHandler:update failed", error);
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
    const onRemoveHandler = () => {
        props.onRemove(props.settingItemId);
    };
    const onMessageBarDismiss = () => {
        setExecutingResult(result => ({ ...result, displayMessage: false }));
    };
    const loadAreas = async (baseUrl: string): Promise<AreaCollection> => {
        const areas = await serviceContext.areaService.getAreasByBaseUrl(baseUrl);
        return areas;
        // const selSettingItem = serviceContext.setting.getSettingItemById(settingItemId);
        // const areas = await serviceContext.areaService.getAreas(selSettingItem);
        // const options = areas.toIDropdownOption(selSettingItem);
        // const defaultArea = serviceContext.areaService.getAreaByOutlookItem(selSettingItem, areas, outlookItem);
        // return {
        //     options,
        //     defaultKey: defaultArea.Id
        // };
    };

    return (
        <div className="ms-Grid">
            {executingResult.displayMessage && <div className="ms-Grid-row"><div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                <MessageBar messageBarType={executingResult.isError ? MessageBarType.error : MessageBarType.success}
                    onDismiss={onMessageBarDismiss}
                    isMultiline={false}>{executingResult.message}</MessageBar>
            </div></div>}
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false && currentSettingItem != null}>
                        {currentSettingItem && <CtrlTextField2
                            label="Setting Name:"
                            name="name"
                            rules={{ required: "Setting Name is required" }}
                            control={control}
                            errors={errors}
                            defaultValue={currentSettingItem.name} />}

                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false && currentSettingItem != null}>
                        {currentSettingItem && <CtrlTextField2
                            label="Azure DevOps URL:"
                            name="baseUrl"
                            rules={{
                                required: "Azure DevOps Url is required",
                                validate: onBaseUrlValidation
                            }}
                            control={control}
                            errors={errors}
                            defaultValue={currentSettingItem.baseUrl} />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false && currentSettingItem != null && currentQueries != null}>
                        {currentSettingItem && <CtrlDropdown
                            label="Task Query:"
                            name="tasksQueryId"
                            rules={{ required: "Task Query is required" }}
                            control={control}
                            errors={errors}
                            options={currentQueries ? currentQueries.toIDropdownOptions() : []}
                            selectedKey={currentSettingItem.tasksQueryId}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false && currentSettingItem != null && currentQueries != null}>
                        {currentSettingItem && <CtrlDropdown
                            label="Issue Query:"
                            name="issuesQueryId"
                            rules={{ required: "Issue Query is required" }}
                            control={control}
                            errors={errors}
                            options={currentQueries ? currentQueries.toIDropdownOptions() : []}
                            selectedKey={currentSettingItem.issuesQueryId}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false && currentSettingItem != null && currentQueries != null}>
                        {currentSettingItem && <CtrlDropdown
                            label="Report Query:"
                            name="reportQueryId"
                            rules={{ required: "Report Query is required" }}
                            control={control}
                            errors={errors}
                            options={currentQueries ? currentQueries.toIDropdownOptions() : []}
                            selectedKey={currentSettingItem.reportQueryId}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false}>
                        {currentSettingItem && <CtrlToggle
                            label="Default setting"
                            name="isDefault"
                            rules={{ required: "Default setting is required" }}
                            control={control}
                            errors={errors}
                            checked={setting.defaultSettingId == props.settingItemId}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={executingResult.isRunning == false}>
                        {currentSettingItem && <CtrlAreaSetting
                            label="Area settings"
                            name="areas"
                            rules={{ required: "Default setting is required" }}
                            control={control}
                            errors={errors}
                            defaultBaseUrl={currentSettingItem.baseUrl}
                            baseUrlWatchName="baseUrl"
                            defaultValue={currentSettingItem.areas}
                            onGetAreas={loadAreas}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Stack horizontal horizontalAlign="space-around" style={{ marginTop: "10px" }}>
                        <PrimaryButton disabled={false} text="Update" onClick={onUpdateHandler} />
                        <DefaultButton disabled={false} text="Remove" onClick={onRemoveHandler} />
                    </Stack>
                </div>
            </div>
        </div>

    );
};

// export default withAITracking(reactPlugin, SettingsView);
export default SettingsView;