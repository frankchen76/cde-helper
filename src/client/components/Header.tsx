import { CommandBar, ICommandBarItemProps } from "@fluentui/react";
import * as React from "react";
import { useContext, useEffect, useState } from "react";
import { ServiceContext } from "../services/SettingService";

export interface IHeaderProps {
    history?: any;
    account?: string;
    //isExistTask: (itemId: any) => boolean;
    //outlookItem: OutlookItem;

};

export const Header = (props: IHeaderProps) => {
    const [cmdBarItems, setCmdbarItems] = useState<ICommandBarItemProps[]>([]);
    const cmdBarOverflowItems: ICommandBarItemProps[] = [
        { key: 'report', text: 'Report', iconProps: { iconName: 'ReportDocument' }, href: "#/tasksreport" },
        { key: 'settings', text: 'Settings', iconProps: { iconName: 'Settings' }, href: "#/setting" }
    ];
    const serviceContext = useContext(ServiceContext);
    const customStyle = {
        "display": "inline",
        "list-style-type": "none"
    };
    const liStyle: React.CSSProperties = {
        "float": "left"
    }

    useEffect(() => {
        let subAllTaskItems: ICommandBarItemProps[] = [];

        // add view tasks menu
        let allTaskItem: ICommandBarItemProps = {
            key: "viewtasks",
            text: "View tasks",
            iconProps: { iconName: 'TaskLogo' },
            //disabled: existingTask != null || props.isDialog,
            href: "#/",
            subMenuProps: {
                items: []
            }
        };
        if (serviceContext.setting.items && Array.isArray(serviceContext.setting.items)) {
            // add all areas under view tasks menu
            serviceContext.setting.items.forEach(settingItem => {
                settingItem.areas.forEach(area => {
                    if (area.enabled) {
                        allTaskItem.subMenuProps.items.push({
                            key: `viewtasks-${settingItem.id}-${area.areaId}`, text: `${settingItem.name}-${area.areaName}`, iconProps: { iconName: 'TaskLogo' }, href: `#/tasksviewbyarea/${settingItem.id}/${area.areaId}`
                        });
                    }
                });
            });
        }

        subAllTaskItems.push(allTaskItem);
        serviceContext.setting.items.forEach(settingItem => {
            subAllTaskItems.push({
                key: `addemailtask-${settingItem.id}`, text: `Add Email Task-${settingItem.name}`, iconProps: { iconName: 'PageAdd' }, href: `#/redirecttaskitem/${settingItem.id}/0`
            });
            subAllTaskItems.push({
                key: `addtask-${settingItem.id}`, text: `Add Task-${settingItem.name}`, iconProps: { iconName: 'PageAdd' }, href: `#/redirecttaskitem/${settingItem.id}/-1`
            });
        });

        let subIssueItems: ICommandBarItemProps[] = [{
            key: "viewissues",
            text: "View issues",
            iconProps: { iconName: 'IssueTracking' },
            href: "#/issuesview"
        }];
        serviceContext.setting.items.forEach(settingItem => {
            subIssueItems.push({
                key: `addissue-${settingItem.id}`, text: `Add Issue-${settingItem.name}`, iconProps: { iconName: 'AddBookmark' }, href: `#/redirectissueitem/${settingItem.id}/0`,
            });
        });

        const items: ICommandBarItemProps[] = [
            {
                key: 'tasks',
                text: 'Tasks',
                iconProps: { iconName: 'TaskLogo' },
                //href: "#/",
                subMenuProps: {
                    items: subAllTaskItems
                }
            },
            { key: 'issues', text: 'Issues', iconProps: { iconName: 'IssueTracking' }, href: "#/issuesview", subMenuProps: { items: subIssueItems } }
        ];

        setCmdbarItems(items);
    }, [serviceContext.setting]);

    return (
        <div className="navbar">
            <CommandBar
                items={cmdBarItems}
                overflowItems={cmdBarOverflowItems}
                ariaLabel="Use left and right arrow keys to navigate between commands"
            />
        </div>
    );
}
//export default withRouter<RouteComponentProps<IHeaderProps>, Header>(Header);
export default Header;