import * as React from "react";
import { useState, useContext } from "react";
import {
    getFocusStyle,
    getTheme,
    IconButton,
    IIconProps,
    mergeStyleSets,
    Separator,
    Link,
    CommandButton,
    IContextualMenuProps,
    ITooltipProps,
    ITooltipHostStyles,
    TooltipHost
} from "@fluentui/react";
import { ServiceContext } from "../services/SettingService";
import { TaskStateComponent } from "./TaskStateComponent";
import { Issue } from "../services/Issue";
import * as _ from "lodash";
import { Common } from "../services/Common";

export interface IIssuesRowProps {
    issue: Issue;
}

export const IssuesRow = (props: IIssuesRowProps) => {
    const [issue, setIssue] = useState(props.issue);
    const serviceContext = useContext(ServiceContext);

    const iconProps: IIconProps = {
        iconName: "IssueTracking",
        style: { fontSize: 15 }
    };
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
    // set tooltip style
    const tooltipStyle = {
        maxHeight: 250,
        overflow: "auto"
    };
    const tooltipProps: ITooltipProps = {
        onRenderContent: () => (
            <div style={tooltipStyle} dangerouslySetInnerHTML={{ __html: `<h2>${issue.title}</h2><hr>${issue.description}` }}>
            </div>
        ),
    };

    const statusHandler = (ev, menuItem) => {
        const { issueService, setting } = serviceContext;
        let newIssue = _.cloneDeep(issue);
        newIssue.state = menuItem.key;
        const settingItem = setting.getSettingItemById(newIssue.settingItemId);
        issueService.updateIssueState(settingItem, newIssue)
            .then(() => {
                setIssue(newIssue);
            });

    };
    const btnContextMenuStyle = {
        height: 20,
        paddingLeft: 0
    };
    const contextMenuItems: IContextualMenuProps = {
        useTargetAsMinWidth: true,
        items: [
            {
                key: "status",
                text: "Status",
                iconProps: { iconName: "StatusCircleRing" },
                subMenuProps: {
                    items: [
                        {
                            key: "ToDo",
                            text: "To Do",
                            iconProps: issue.state == "To Do" ? { iconName: "StatusCircleCheckmark" } : {},
                            onClick: statusHandler
                        },
                        {
                            key: "Doing",
                            text: "Doing",
                            iconProps: issue.state == "Doing" ? { iconName: "StatusCircleCheckmark" } : {},
                            onClick: statusHandler
                        },
                        {
                            key: "Done",
                            text: "Done",
                            iconProps: issue.state == "Done" ? { iconName: "StatusCircleCheckmark" } : {},
                            onClick: statusHandler
                        }
                    ]
                }
            }
        ]
    };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

    return issue ? (
        <div className={`${classNames.itemCell} ms-Grid`} dir="ltr">
            <div className="ms-Grid-row" style={{ fontSize: "14px" }}>
                <div className="ms-Grid-col ms-sm2">
                    <IconButton iconProps={iconProps} title="Issue" />
                </div>
                <div className="ms-Grid-col ms-sm8">
                    <TooltipHost tooltipProps={tooltipProps}
                        styles={hostStyles}
                        closeDelay={500}
                        calloutProps={{ gapSpace: 0 }}>
                        <Link href={`#/issueitem/${issue.settingItemId}/${issue.id}`} >{`${issue.title}`}</Link>
                    </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm2">
                    <TaskStateComponent state={issue.state} />
                </div>
            </div>
            <div className="ms-Grid-row" style={{ fontSize: "11px" }}>
                <div className="ms-Grid-col ms-sm2">
                    <span style={{ fontWeight: "bold" }}>ID</span>
                </div>
                <div className="ms-Grid-col ms-sm3">
                    <span style={{ fontWeight: "bold" }}>Created: </span>
                </div>
                <div className="ms-Grid-col ms-sm5">
                    <span style={{ fontWeight: "bold" }}>Axis#: </span>
                </div>
                <div className="ms-Grid-col ms-sm2" style={{ paddingLeft: 0 }}>
                    <CommandButton style={btnContextMenuStyle} text="" menuProps={contextMenuItems} />
                </div>
            </div>
            <div className="ms-Grid-row" style={{ fontSize: "11px" }}>
                <div className="ms-Grid-col ms-sm2">
                    {issue.id}
                </div>
                <div className="ms-Grid-col ms-sm3">
                    {Common.dateToDurationString(issue.createdDate)}
                </div>
                <div className="ms-Grid-col ms-sm7 divIssueItem">
                    {issue.axisCode}
                </div>
            </div>
            {/* <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm10 ms-smPush2">
                    <span style={{ fontWeight: "bold" }}>Created: </span>{issue.createdDateDuration} days ago
                </div>
            </div> */}
            <Separator styles={{ root: { padding: 0 } }} />
        </div>
    ) : null;
}