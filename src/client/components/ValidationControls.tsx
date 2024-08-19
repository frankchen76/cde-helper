import { Checkbox, ComboBox, DatePicker, Dropdown, IComboBoxOption, IDropdownOption, IDropdownStyles, ITag, Position, PrimaryButton, Shimmer, SpinButton, Spinner, TextField, Toggle } from "@fluentui/react";
import * as _ from "lodash";
import * as React from "react";
import { useEffect, useState } from "react";
import { Control, Controller, useController, useWatch } from "react-hook-form";
import { AreaCollection } from "../services/Area";
import { ISettingAreaItem, ISettingItem } from "../services/SettingService";
import { IQuillComponentProps, QuillEditor } from "./QuillEditor";
import { ShimmerCtrl } from "./TaskItemShimmer";
import { TaskItemTagPicker } from "./TaskItemTagPicker";

export interface ICtrlProps {
    label: string;
    name: string;
    rules: any;
    control: Control<any>;
    errors: any;
    required?: boolean;
};
export interface ICtrlToggleProps extends ICtrlProps {
    checked: boolean;
};
export const CtrlToggle = (props: ICtrlToggleProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    const [checked, setChecked] = useState<boolean>(props.checked);
    useEffect(() => {
        //field.value = props.selectedKey;
        // if (props.setValue) {
        //     props.setValue(props.name, props.selectedKey);
        // }
        field.onChange(props.checked);
    }, [props.checked]);
    return (
        <Controller
            name={props.name}
            control={props.control}
            rules={props.rules}
            defaultValue={props.checked}
            render={({ field }) => {
                //console.log("toggle value:", field.value, "defaultValue:", props.checked);
                return (
                    <div>
                        <Toggle label="Default Setting"
                            onText="Yes"
                            offText="No"
                            checked={checked}
                            onChange={(event, checked) => {
                                setChecked(checked);
                                field.onChange(checked.toString());
                            }}
                        />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                    </div>)
            }}
        />

    );
};
export interface ICtrlTextFieldProps extends ICtrlProps {
    defaultValue: string;
    multiline?: boolean;
};
export const CtrlTextField = (props: ICtrlTextFieldProps) => {
    return (
        <Controller
            name={props.name}
            control={props.control}
            rules={props.rules}
            defaultValue={props.defaultValue}
            render={({ field }) => {
                //console.log("value:", field.value, "defaultValue:", props.defaultValue);
                return (
                    <div>
                        <TextField required={props.required} label={props.label} multiline={props.multiline} {...field} />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                        <span className="formError">{props.errors[props.name]?.type === 'validate' && props.errors[props.name].message}</span>
                    </div>)
            }}
        />

    );
};
export const CtrlTextField2 = (props: ICtrlTextFieldProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    useEffect(() => {
        //field.value = props.selectedKey;
        // if (props.setValue) {
        //     props.setValue(props.name, props.selectedKey);
        // }
        field.onChange(props.defaultValue);
    }, [props.defaultValue]);
    return (
        <Controller
            name={props.name}
            control={props.control}
            rules={props.rules}
            defaultValue={props.defaultValue}
            render={({ field }) => {
                //console.log("value:", field.value, "defaultValue:", props.defaultValue);
                return (
                    <div>
                        <TextField required={props.required} label={props.label} {...field} />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                        <span className="formError">{props.errors[props.name]?.type === 'validate' && props.errors[props.name].message}</span>
                    </div>)
            }}
        />

    );
};
export interface ICtrlDatePickerProps extends ICtrlProps {
    defaultValue?: Date;
    minDate?: Date;
    maxDate?: Date;
};
export const CtrlDatePicker = (props: ICtrlDatePickerProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    useEffect(() => {
        field.onChange(props.defaultValue);
    }, [props.defaultValue]);
    return (
        <Controller
            name={props.name}
            control={props.control}
            rules={props.rules}
            defaultValue={props.defaultValue}
            render={({ field }) => {
                //console.log("CtrlDatePicker value:", field.value, "defaultValue:", props.defaultValue);
                return (
                    <div>
                        <DatePicker
                            isRequired={props.required}
                            today={new Date()}
                            label={props.label}
                            placeholder="Select a date..."
                            ariaLabel="Select a date"
                            minDate={props.minDate}
                            maxDate={props.maxDate}
                            value={field.value}
                            onSelectDate={(date): void => field.onChange(date)}
                        />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                        <span className="formError">{props.errors[props.name]?.type === 'validate' && props.errors[props.name].message}</span>
                    </div>)
            }}
        />

    );
};

export interface ICtrlDropdownProps extends ICtrlProps {
    options: IDropdownOption[];
    selectedKey: string | number;
    disabled?: boolean;
};
export const CtrlDropdown = (props: ICtrlDropdownProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    useEffect(() => {
        //field.value = props.selectedKey;
        // if (props.setValue) {
        //     props.setValue(props.name, props.selectedKey);
        // }
        field.onChange(props.selectedKey);
    }, [props.selectedKey]);
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            defaultValue={props.selectedKey}
            control={props.control}
            render={({ field }) => {
                //console.log("field name:", field.name, "value:", props.selectedKey, "options:", props.options);
                return (
                    <div>
                        <Dropdown
                            required={props.required}
                            label={props.label}
                            {...field}
                            options={props.options}
                            onChange={(e, option): void => field.onChange(option.key)}
                            selectedKey={field.value}
                            disabled={props.disabled}
                        />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                    </div>
                );
            }}
        />
    );
};
export interface ICtrlWatchDropdownOptions {
    options: IDropdownOption[];
    defaultKey: number | string;
}
export interface ICtrlWatchDropdownProps extends ICtrlProps {
    watchedName1?: string;
    watchedName2?: string;
    defaultValue?: string | number;
    disabled?: boolean;
    onGetOptions: (...param: string[]) => Promise<ICtrlWatchDropdownOptions>;
};
export const CtrlWatchDropdown = (props: ICtrlWatchDropdownProps) => {
    //const service = useContext(AzureDevOpsServiceContext);
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    const watchField1 = props.watchedName1 == null ? null : useWatch({
        control: props.control,
        name: props.watchedName1
    });
    const watchField2 = props.watchedName2 == null ? null : useWatch({
        control: props.control,
        name: props.watchedName2
    });
    const effectArray = [];
    if (watchField1 != null) effectArray.push(watchField1);
    if (watchField2 != null) effectArray.push(watchField2);

    const [loaded, setLoaded] = useState<boolean>(false);
    const [options, setOptions] = useState<IDropdownOption[]>();
    const [selectedKey, setSelectedKey] = useState<string | number>(props.defaultValue);
    //console.log(props.name);

    useEffect(() => {
        const loadOptions = async () => {
            setLoaded(false);
            const ddnOptions = await props.onGetOptions(watchField1, watchField2);
            if (ddnOptions) {
                const existOption = _.find(ddnOptions.options, { key: props.defaultValue }) != null;
                setOptions(ddnOptions.options);
                if (props.defaultValue && existOption) {
                    setSelectedKey(props.defaultValue);
                    field.onChange(props.defaultValue);
                } else {
                    setSelectedKey(ddnOptions.defaultKey);
                    field.onChange(ddnOptions.defaultKey);
                }
            }
            setLoaded(true);
        };

        loadOptions();
    }, [watchField1, watchField2]);
    const dropdownStyles: Partial<IDropdownStyles> = {
        dropdownOptionText: { overflow: 'visible', whiteSpace: 'normal' },
        dropdownItem: { height: 'auto' },
    };
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            defaultValue={selectedKey}
            control={props.control}
            render={({ field }) => {
                //console.log("field name:", field.name, "value:", props.selectedKey, "options:", props.options);
                return (
                    <div>
                        <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={loaded}>
                            <Dropdown
                                label={props.label}
                                required={props.required}
                                options={options ? options : []}
                                onChange={(e, option): void => field.onChange(option.key)}
                                selectedKey={field.value}
                                disabled={props.disabled}
                                styles={dropdownStyles}
                            />
                            <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                        </Shimmer>
                    </div>
                );
            }}
        />
    );

};
export interface ICtrlComboBoxProps extends ICtrlProps {
    defaultValue: string | number;
    disabled?: boolean;
    watchedName1?: string;
    watchedName2?: string;
    onGetOptions: (...param: string[]) => Promise<ICtrlWatchDropdownOptions>;

}
export const CtrlComboBox = (props: ICtrlComboBoxProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    const watchField1 = props.watchedName1 == null ? null : useWatch({
        control: props.control,
        name: props.watchedName1
    });
    const watchField2 = props.watchedName2 == null ? null : useWatch({
        control: props.control,
        name: props.watchedName2
    });
    const [loaded, setLoaded] = useState<boolean>(false);
    const [options, setOptions] = useState<IDropdownOption[]>();
    const [selectedKey, setSelectedKey] = useState<string | number>(props.defaultValue);

    useEffect(() => {
        const loadOptions = async () => {
            setLoaded(false);
            const ddnOptions = await props.onGetOptions(watchField1, watchField2);
            if (ddnOptions) {
                const existOption = _.find(ddnOptions.options, { key: props.defaultValue }) != null;
                setOptions(ddnOptions.options);
                if (props.defaultValue && existOption) {
                    setSelectedKey(props.defaultValue);
                    field.onChange(props.defaultValue);
                } else {
                    setSelectedKey(ddnOptions.defaultKey);
                    field.onChange(ddnOptions.defaultKey);
                }
            }
            setLoaded(true);
        };

        loadOptions();
        // field.onChange(props.defaultValue);
    }, [watchField1, watchField2]);
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            defaultValue={props.defaultValue}
            control={props.control}
            render={({ field }) => {
                //console.log("field name:", field.name, "value:", props.selectedKey, "options:", props.options);
                return (
                    <div>
                        <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={options != null}>
                            {options && <ComboBox label={props.label}
                                required={props.required}
                                text={field.value}
                                selectedKey={field.value}
                                options={options}
                                allowFreeform={true}
                                onChange={(event, option, index, val): void => field.onChange(val ? val : option.key.toString())}
                            />}
                            {/* <Dropdown
                            label={props.label}
                            {...field}
                            options={props.options}
                            onChange={(e, option): void => field.onChange(option.key)}
                            selectedKey={field.value}
                            disabled={props.disabled}
                        /> */}
                            <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                        </Shimmer>
                    </div>
                );
            }}
        />
    );
};
export interface ICtrlQuillProps extends IQuillComponentProps, ICtrlProps {
    watchedTitleName?: string;
};
export const CtrlQuill = (props: ICtrlQuillProps) => {
    const watchTitleField = props.watchedTitleName == null ? null : useWatch({
        control: props.control,
        name: props.watchedTitleName
    });
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            defaultValue={props.html}
            control={props.control}
            render={({ field }) => {
                return (
                    <div>
                        <QuillEditor html={field.value}
                            required={props.required}
                            label={props.label}
                            placeholder={props.placeholder}
                            outlookItem={props.outlookItem}
                            taskTitle={watchTitleField}
                            onChange={(content: string): void => {
                                field.onChange(content);
                            }} />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                    </div>
                );
            }}
        />

    );
};
export interface ICtrlSpinButtonProps extends ICtrlProps {
    min: number;
    max: number;
    step: number;
    defaultValue: string;
};
export const CtrlSpinButton = (props: ICtrlSpinButtonProps) => {
    const [num, setNum] = useState<number>(+(props.defaultValue));
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    useEffect(() => {
        field.onChange(props.defaultValue);
    }, []);
    const onCompletedHourIncreased = (val: string): string => {
        let existNum: number = +val;
        if (existNum + props.step <= props.max) {
            existNum += props.step;
            setNum(existNum);
            field.onChange(existNum);
        }
        return existNum.toString();
    };
    const onCompletedHourDecreased = (val: string): string => {
        let existNum: number = +val;
        if (existNum - props.step >= props.min) {
            existNum -= props.step;
            setNum(existNum);
            field.onChange(existNum);
        }
        return existNum.toString();
    };
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            defaultValue={props.defaultValue}
            control={props.control}
            render={({ field }) => {
                //console.log("field name:", field.name, "value:", props.defaultValue);
                return (
                    <div>
                        <SpinButton
                            // value={props.defaultValue}
                            defaultValue={props.defaultValue}
                            label={props.label}
                            labelPosition={Position.top}
                            min={props.min}
                            max={props.max}
                            step={props.step}
                            // incrementButtonAriaLabel={'Increase value by props.step'}
                            onIncrement={onCompletedHourIncreased}
                            onDecrement={onCompletedHourDecreased}
                        // decrementButtonAriaLabel={'Decrease value by props.step'} 
                        />
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                    </div>
                );
            }}
        />

    );
};
export interface ICtrlTaskTagPickerProps extends ICtrlProps {
    defaultValue: string;
    //settingItem: ISettingItem;
    watchedNameSettingItem: string;
    onGetSettingItem: (settingItemId: string) => Promise<ISettingItem>;
};
export const CtrlTaskTagPicker = (props: ICtrlTaskTagPickerProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    const watchFieldSettingItem = useWatch({
        control: props.control,
        name: props.watchedNameSettingItem
    });
    const [settingItem, setSettingItem] = useState<ISettingItem>();

    useEffect(() => {
        const loadSettingItem = async () => {
            const sItem = await props.onGetSettingItem(watchFieldSettingItem);
            if (sItem) {
                setSettingItem(sItem);
            }
            field.onChange(props.defaultValue);
        };

        loadSettingItem();
    }, [watchFieldSettingItem]);

    const onTagPickerChange = (items?: ITag[]): void => {
        const tags = items ? items.map(t => t.name).join(",") : "";
        field.onChange(tags);
    }
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            defaultValue={props.defaultValue}
            control={props.control}
            render={({ field }) => {
                //console.log("field name:", field.name, "value:", props.defaultValue);
                return (
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={settingItem != null}>
                        {settingItem && <TaskItemTagPicker label={props.label}
                            required={props.required}
                            defaultSelectedTags={props.defaultValue}
                            onTagPickerChange={onTagPickerChange}
                            settingItem={settingItem} />}
                        <span className="formError">{props.errors[props.name]?.type === 'required' && props.errors[props.name].message}</span>
                    </Shimmer>
                );
            }}
        />

    );
};
export interface ICtrlAreaSettingProps extends ICtrlProps {
    baseUrlWatchName: string;
    defaultBaseUrl: string;
    defaultValue: ISettingAreaItem[];
    onGetAreas: (baseUrl: string) => Promise<AreaCollection>;
};
export const CtrlAreaSetting = (props: ICtrlAreaSettingProps) => {
    const { field } = useController({
        control: props.control,
        name: props.name
    });
    const watchFieldBaseUrl = useWatch({
        control: props.control,
        name: props.baseUrlWatchName
    });
    const [loaded, setLoaded] = useState<boolean>(false);
    const [currentSettingAreaItems, setCurrentSettingAreaItems] = useState<ISettingAreaItem[]>();

    // load 
    useEffect(() => {
        loadAreasCtrl(props.defaultBaseUrl);

    }, []);
    useEffect(() => {
        loadAreasCtrl(watchFieldBaseUrl);
    }, [watchFieldBaseUrl]);

    const loadAreasCtrl = async (baseUrl: string): Promise<void> => {
        if (baseUrl != null && baseUrl != "") {
            setLoaded(false);
            const allAreas = await props.onGetAreas(baseUrl);
            let newAreaSettings: ISettingAreaItem[] = [];
            allAreas.items.forEach(area => {
                let existAreaSetting = _.find(props.defaultValue, { areaId: area.Id });
                if (existAreaSetting == null) {
                    existAreaSetting = {
                        areaId: area.Id,
                        areaName: area.Name,
                        enabled: true,
                        emailDomains: []
                    };
                }
                newAreaSettings.push(existAreaSetting);
            });
            setCurrentSettingAreaItems(newAreaSettings);
            field.onChange(newAreaSettings);
            setLoaded(true);
        }
    };

    const onAreaEnabledChanged = (areaSetting: ISettingAreaItem, ev, checked: boolean) => {
        //console.log(areaSetting);
        const newAreItems = _.cloneDeep(currentSettingAreaItems);
        const existIndex = _.findIndex(newAreItems, { areaId: areaSetting.areaId });
        if (existIndex != -1) {
            newAreItems[existIndex].enabled = checked;
            setCurrentSettingAreaItems(newAreItems);
            field.onChange(newAreItems);
        }
    };

    const onEmailDomainsChanged = (areaSetting: ISettingAreaItem, ev, newValue: string) => {
        const newAreItems = _.cloneDeep(currentSettingAreaItems);
        const existIndex = _.findIndex(newAreItems, { areaId: areaSetting.areaId });
        if (existIndex != -1) {
            let emailDomains: string[] = [];
            if (newValue.indexOf(";") == -1) {
                emailDomains.push(newValue);
            } else {
                emailDomains = newValue.split(";");
                // newValue.split(";").forEach(d => {
                //     if (d != "") emailDomains.push(d);
                // });
            }
            newAreItems[existIndex].emailDomains = emailDomains;
            setCurrentSettingAreaItems(newAreItems);
            field.onChange(newAreItems);
        }

    };

    const renderAreaSetting = (areaSettings: ISettingAreaItem[]) => {
        return areaSettings.map((areaSetting, index) => {
            return (
                <div key={`divrow_${index}`} className="ms-Grid-row" style={{ marginBottom: "5px" }}>
                    <div key={`divrowEnabled_${index}`} className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
                        <Checkbox key={`chkEnabled_${index}`} label={areaSetting.areaName} checked={areaSetting.enabled} onChange={onAreaEnabledChanged.bind(this, areaSetting)} />
                    </div>
                    <div key={`divrowEamilDomains_${index}`} className="ms-Grid-col ms-sm8 ms-md8 ms-lg8">
                        <TextField key={`txtEmailDomains_${index}`} value={areaSetting.emailDomains?.join(";")} onChange={onEmailDomainsChanged.bind(this, areaSetting)} />
                    </div>
                </div>
            );
        });
    };
    return (
        <Controller
            name={props.name}
            rules={props.rules}
            //defaultValue={selectedKey}
            control={props.control}
            render={({ field }) => {
                //console.log("field name:", field.name, "value:", props.selectedKey, "options:", props.options);
                return (
                    <Shimmer customElementsGroup={ShimmerCtrl()} isDataLoaded={loaded}>
                        <div className="ms-Grid-row" style={{ marginBottom: "5px" }}>
                            <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
                                Enabled
                            </div>
                            <div className="ms-Grid-col ms-sm8 ms-md8 ms-lg8">
                                Email domains
                            </div>
                        </div>
                        {currentSettingAreaItems && renderAreaSetting(currentSettingAreaItems)}
                    </Shimmer>
                );
            }}
        />
    );

};

export interface ISpinnerButtonProps {
    isRunning?: boolean;
    text: string;
    onClick: () => void;
}
export const SpinnerButton = (props: ISpinnerButtonProps) => {
    // const [isRunning, setIsRunning] = useState<boolean>(false);
    // const onClickHandler = async () => {
    //     setIsRunning(true);
    //     await props.onClick();
    //     setIsRunning(false);
    // }
    return (
        <div style={{ position: 'relative', width: 80 }}>
            <PrimaryButton text={props.text} onClick={props.onClick} disabled={props.isRunning} />
            {props.isRunning && <Spinner styles={{ root: { position: 'absolute', top: 6 } }} />}
        </div>
    );
}

