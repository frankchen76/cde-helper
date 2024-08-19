import {
    DefaultButton,
    IComboBoxOption,
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
import { ISettingItem, ServiceContext } from "../services/SettingService";
import { cloneDeep } from "lodash";
import { Area, AreaCollection } from "../services/Area";
// import { withAITracking } from '@microsoft/applicationinsights-react-js';
// import { reactPlugin } from '../services/AppInsights';
import { IssueItemShimmer } from "./IssueItemShimmer";
import { useContext, useEffect, useState } from "react";
import { IssueFormModeEnum, Common, StateEnum, ExecutingResult } from "../services/Common";
import { useForm } from "react-hook-form";
import { CtrlComboBox, CtrlDropdown, CtrlQuill, CtrlTextField, CtrlWatchDropdown, ICtrlWatchDropdownOptions } from "./ValidationControls";
import { ShimmerButtons, ShimmerCtrl } from "./TaskItemShimmer";

type IssueFormInput = {
    title: string;
    settingItemId: string;
    areaId: number;
    stateId: string;
    axisCode: string;
    description: string;
};
export interface IIssueItemViewProps {
}
const IssueItemView = (props: IIssueItemViewProps) => {
    const serviceContext = useContext(ServiceContext);
    const currentUrl = new URL(window.location.href);
    const hostInfo = currentUrl.searchParams.get("_host_Info");
    const isDialog = hostInfo && hostInfo.indexOf("isDialog") != -1;

    // const id: number = +props.routeProps["match"]["params"]["id"];
    // const settingItemId: string = props.routeProps["match"]["params"]["settingItemId"];
    const para = useParams();
    const history = useHistory();
    const id: number = +para["id"];
    const settingItemId: string = para["settingItemId"];

    const [settingItem, setSettingItem] = useState<ISettingItem>(serviceContext.setting.getSettingItemById(settingItemId));
    const [formMode, setFormMode] = useState<IssueFormModeEnum>(Common.getIssueFormModeFromId(id));

    const [currentIssue, setCurrentIssue] = useState<Issue>();
    const [allAxisCodeOptions, setAllAxisCodeOptions] = useState<IComboBoxOption[]>();
    const [executingResult, setExecutingResult] = useState<ExecutingResult>(ExecutingResult.createInstance());

    const stateOptions = Common.getStateOptions();

    // validation
    const { control, watch, formState: { errors }, handleSubmit, setValue } = useForm<IssueFormInput>();

    useEffect(() => {
        const loadIssues = async () => {
            try {
                setExecutingResult(ExecutingResult.start());
                const { areaService, issueService } = serviceContext;

                //const axisCodeOptions = issues ? issues.getIssueAxisCodeIDropdownOptionsByArea(selArea.Id, false) : [];
                //const axisCode = axisCodeOptions ? axisCodeOptions[0].key.toString() : "";

                //get current iternation id
                const currentIternationId = await areaService.getCurrentIterationId(settingItem);
                if (currentIternationId == 0) {
                    console.error("Cannot get the current Iternation");
                }

                let issue: Issue = null;

                if (id == 0) {
                    //init email/event task

                    issue = new Issue(0,
                        "",
                        undefined,
                        undefined,
                        undefined,
                        currentIternationId,//iterationId
                        StateEnum.ToDo,
                        "",
                        "",
                        undefined,
                        new Date(),
                        settingItem.id);//axisCode
                } else {
                    issue = await issueService.getIssueById(settingItem, id);
                    //selArea = areas.getAreaById(issue.areaId);
                }
                setCurrentIssue(issue);
                //setAllAxisCodeOptions(axisCodeOptions);
                setExecutingResult(ExecutingResult.complete());

            } catch (error) {
                console.log(error);
                setExecutingResult(ExecutingResult.complete(true, Common.getErrorMessage(error), true));
            }

        };

        loadIssues();
    }, []);

    const onAddHandler = async (): Promise<void> => {
        const { issueService, setting } = serviceContext;

        handleSubmit(
            async (data) => {
                console.log("validation successed:", data);
                try {
                    setExecutingResult(ExecutingResult.start());
                    let issueValue = cloneDeep(currentIssue);
                    issueValue.title = data.title;
                    issueValue.settingItemId = data.settingItemId;
                    issueValue.areaId = data.areaId;
                    issueValue.state = data.stateId;
                    issueValue.axisCode = data.axisCode;
                    issueValue.description = data.description;
                    // Update Azure DevOps
                    const newIssue = await issueService.updateIssue(setting.upn, settingItem, issueValue, currentIssue);

                    setCurrentIssue(issueValue);
                    setFormMode(IssueFormModeEnum.UpdateIssue);
                    setExecutingResult(ExecutingResult.complete(true, "Issue was Updated"));
                } catch (error) {
                    console.log(error);
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
        history.push("/issuesview");
    };
    const onMessageBarDismiss = () => {
        setExecutingResult(result => ({ ...result, displayMessage: false }));
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
            ret = {
                options,
                defaultKey: options && options.length > 0 ? options[0].key : undefined
            };
        }
        return ret;
    };
    const loadAxisCodes = async (settingItemId: string, areaId: string): Promise<ICtrlWatchDropdownOptions> => {
        let ret: ICtrlWatchDropdownOptions = null;
        if (settingItemId != undefined && areaId != undefined) {
            const selSettingItem = serviceContext.setting.getSettingItemById(settingItemId);
            const selArea = await serviceContext.areaService.getAreaById(selSettingItem, +areaId);
            const allIssues = await serviceContext.issueService.getIssuesByArea(selSettingItem, selArea);
            const options = allIssues.getIssueAxisCodeIDropdownOptionsByArea(selArea.Id, false);
            ret = {
                options,
                defaultKey: options && options.length > 0 ? options[0].key : undefined
            };
        }
        return ret;
    };

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>{formMode == IssueFormModeEnum.UpdateIssue ? "Update issue" : "Add issue"}</h2>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    {executingResult.displayMessage &&
                        <MessageBar messageBarType={executingResult.isError ? MessageBarType.error : MessageBarType.success}
                            onDismiss={onMessageBarDismiss}
                            isMultiline={false}>{executingResult.message}</MessageBar>
                    }
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={currentIssue != null}>
                        {currentIssue && <CtrlTextField
                            label="Title:"
                            name="title"
                            rules={{ required: "Title is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={currentIssue.title} />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={currentIssue != null}>
                        {currentIssue && <CtrlWatchDropdown
                            label="Setting:"
                            name="settingItemId"
                            rules={{ required: "Setting is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={settingItem.id}
                            disabled={formMode == IssueFormModeEnum.UpdateIssue}
                            onGetOptions={loadSettings}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={currentIssue != null}>
                        {currentIssue && <CtrlWatchDropdown
                            label="Area:"
                            name="areaId"
                            rules={{ required: "Area is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={currentIssue ? currentIssue.areaId : undefined}
                            onGetOptions={loadAreas}
                            watchedName1="settingItemId"
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={currentIssue != null}>
                        {currentIssue && <CtrlDropdown
                            label="State:"
                            name="stateId"
                            rules={{ required: "State is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            options={stateOptions}
                            selectedKey={currentIssue.state}
                        />}
                    </Shimmer>
                </div>
                <div className="ms-Grid-col ms-sm8 ms-md8 ms-lg8" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={currentIssue != null}>
                        {currentIssue && <CtrlComboBox
                            label="Axis Code:"
                            name="axisCode"
                            rules={{ required: "Axis code is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            defaultValue={currentIssue.axisCode}
                            onGetOptions={loadAxisCodes}
                            watchedName1="settingItemId"
                            watchedName2="areaId"
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={currentIssue != null}>
                        {currentIssue && <CtrlQuill
                            label="Description:"
                            name="description"
                            rules={{ required: "Description is required" }}
                            control={control}
                            errors={errors}
                            required={true}
                            html={currentIssue.description}
                            placeholder="Enter description."
                        />}
                    </Shimmer>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" >
                    <Shimmer customElementsGroup={ShimmerButtons()} isDataLoaded={currentIssue != null}>
                        <Stack horizontal horizontalAlign="space-around" style={{ marginTop: "10px" }}>
                            <PrimaryButton disabled={executingResult.isRunning} text={formMode == IssueFormModeEnum.UpdateIssue ? "Update" : "Add"} onClick={onAddHandler} />
                            <DefaultButton disabled={executingResult.isRunning} text="Cancel" onClick={onCancelHandler} />
                        </Stack>
                    </Shimmer>
                </div>
            </div>
            {
                // <Stack tokens={Common.CONTAINER_STACK_TOKENS}>
                //     {/* <TextField required label="Title" value={this.state.issue.title} title={this.state.issue.title} onChange={this._onTitleChanged.bind(this)} /> */}
                //     <CtrlTextField
                //         label="Title:"
                //         name="title"
                //         rules={{ required: "Title is required" }}
                //         control={control}
                //         errors={errors}
                //         defaultValue={currentIssue.title} />

                //     <Stack horizontal tokens={Common.CONTAINER_STACK_TOKENS}>
                //         <Stack.Item grow>
                //             <CtrlWatchDropdown
                //                 label="Setting:"
                //                 name="settingItemId"
                //                 rules={{ required: "Setting is required" }}
                //                 control={control}
                //                 errors={errors}
                //                 defaultValue={settingItem.id}
                //                 disabled={formMode == IssueFormModeEnum.UpdateIssue}
                //                 onGetOptions={loadSettings}
                //             />
                //         </Stack.Item>

                //         <Stack.Item grow>
                //             {currentIssue && <CtrlWatchDropdown
                //                 label="Area:"
                //                 name="areaId"
                //                 rules={{ required: "Area is required" }}
                //                 control={control}
                //                 errors={errors}
                //                 defaultValue={currentIssue ? currentIssue.areaId : undefined}
                //                 onGetOptions={loadAreas}
                //                 watchedName1="settingItemId"
                //             />}
                //         </Stack.Item>
                //     </Stack>
                //     {/* <Dropdown required label="State" options={stateOptions} selectedKey={this.state.issue.state} onChange={this._onStateChanged.bind(this)} /> */}
                //     <CtrlDropdown
                //         label="State:"
                //         name="stateId"
                //         rules={{ required: "State is required" }}
                //         control={control}
                //         errors={errors}
                //         options={stateOptions}
                //         selectedKey={currentIssue.state}
                //     />
                //     {/* <ComboBox label="Axis Code"
                //         text={this.state.issue.axisCode}
                //         selectedKey={this.state.issue.axisCode}
                //         options={axisCodeOptions}
                //         allowFreeform={true}
                //         onChange={this._onAxisCodeChanged.bind(this)}
                //     /> */}
                //     <CtrlComboBox
                //         label="Axis Code:"
                //         name="axisCode"
                //         rules={{ required: "Axis code is required" }}
                //         control={control}
                //         errors={errors}
                //         defaultValue={currentIssue.axisCode}
                //         onGetOptions={loadAxisCodes}
                //         watchedName1="settingItemId"
                //         watchedName2="areaId"
                //     />
                //     {/* <TextField required label="Axis Code" value={this.state.issue.axisCode} title={this.state.issue.axisCode} onChange={this._onAxisCodeChanged.bind(this)} /> */}
                //     {/* <Label required>Description</Label>
                //     <ReactQuill value={this.state.issue.description}
                //         onChange={this._onQuillDescriptionChanged.bind(this)}
                //         //onBlur={this._onQuillDescriptionBlur.bind(this)}
                //         className="myQuill"
                //         modules={modules}
                //         formats={formats} /> */}
                //     <CtrlQuill
                //         label="Description:"
                //         name="description"
                //         rules={{ required: "Description is required" }}
                //         control={control}
                //         errors={errors}
                //         html={currentIssue.description}
                //         placeholder="Enter description."
                //     />

                //     {/* </ReactQuill> */}
                //     <Stack horizontal horizontalAlign="space-around">
                //         <PrimaryButton disabled={executingResult.isRunning} text={formMode == IssueFormModeEnum.UpdateIssue ? "Update" : "Add"} onClick={onAddHandler} />
                //         <DefaultButton disabled={executingResult.isRunning} text="Cancel" onClick={onCancelHandler} />
                //     </Stack>
                // </Stack>
            }
        </div>
    );
};

// export default withAITracking(reactPlugin, IssueItemView);
export default IssueItemView;