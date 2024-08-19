import { ITag, Label, TagPicker } from "@fluentui/react";
import * as React from "react";
import { useContext, useEffect, useState } from "react";
import { ISettingItem, ServiceContext } from "../services/SettingService";
import { TagCollection } from "../services/Tag";

export interface ITaskItemTagPickerProps {
    label: string;
    settingItem: ISettingItem;
    onTagPickerChange: (items?: ITag[]) => void;
    defaultSelectedTags: string;
    required?: boolean;
}
export const TaskItemTagPicker = (props: ITaskItemTagPickerProps) => {
    const serviceContext = useContext(ServiceContext);
    const [tags, setTags] = useState<TagCollection>();
    const [defaultItems, setDefaultItems] = useState<ITag[]>();
    useEffect(() => {
        const loadTags = async () => {
            const { tagService } = serviceContext
            const result = await tagService.getTags(props.settingItem);
            const dItems = result.getITagsFromTagString(props.defaultSelectedTags);
            setTags(result);
            setDefaultItems(dItems);
            console.log("set default tag", dItems);
        };
        loadTags();
    }, [props.settingItem, props.defaultSelectedTags]);
    const listContainsTagList = (tag: ITag, tagList?: ITag[]) => {
        if (!tagList || !tagList.length || tagList.length === 0) {
            return false;
        }
        return tagList.some(compareTag => compareTag.key.toString().toLowerCase() === tag.key.toString().toLowerCase());
    };
    const filterSuggestedTags = (filterText: string, tagList: ITag[]): ITag[] => {
        const iTags = tags.toITags();
        // return filterText
        //     ? tags.filter(
        //         tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 && !this._listContainsTagList(tag, tagList),
        //     )
        //     : [];
        let ret: ITag[] = filterText
            ? iTags.filter(
                tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 && !listContainsTagList(tag, tagList),
            )
            : [];
        if (ret.length == 0) {
            const newTag = {
                name: filterText,
                key: filterText
            };
            if (!listContainsTagList(newTag, tagList))
                ret.push(newTag);
        }
        return ret;
    };
    return (
        <div>
            <Label required={props.required}>{props.label}</Label>
            {defaultItems && <TagPicker
                removeButtonAriaLabel="Remove"
                onResolveSuggestions={filterSuggestedTags}
                getTextFromItem={(item: ITag) => item.name}
                pickerSuggestionsProps={{ suggestionsHeaderText: 'Suggested tags', noResultsFoundText: 'No tag found' }}
                //itemLimit={5}
                onChange={props.onTagPickerChange}
                defaultSelectedItems={defaultItems}
            />}
        </div>

    );
};