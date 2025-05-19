import * as React from "react";
import { useState, useEffect, useContext } from "react";
import { getTheme } from "@fluentui/react/lib/Styling";
import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { GroupedList, IGroup } from "@fluentui/react/lib/GroupedList";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { SelectionMode } from "@fluentui/react/lib/Selection";

import { RouteComponentProps } from "react-router-dom";
import { ServiceContext } from "../services/SettingService";
import { OutlookItem } from "../services/OutlookItem";
import { Task } from "../services/Task";
// import { withAITracking } from '@microsoft/applicationinsights-react-js';
// import { reactPlugin, appInsights } from '../services/AppInsights';
import { Issue, IssueCollection } from "../services/Issue";
import { IssuesViewHeader } from "./IssuesViewHeader";
import cloneDeep = require("lodash/cloneDeep")
import { IssuesRow } from "./IssuesRow";

export interface IIssuesViewProps {
    routeProps: RouteComponentProps;
}

const IssuesView = (props: IIssuesViewProps) => {
    const currentUrl = new URL(window.location.href);
    const hostInfo = currentUrl.searchParams.get("_host_Info");

    const [issues, setIssues] = useState<IssueCollection>();
    const [groups, setGroups] = useState<IGroup[]>();
    const [error, setError] = useState<string>();
    const [loading, setLoading] = useState<boolean>(false);
    const [outlookItem, setOutlookItem] = useState<OutlookItem>();
    const [isDialog, setIsDialog] = useState<boolean>(hostInfo && hostInfo.indexOf("isDialog") != -1);
    const serviceContext = useContext(ServiceContext);

    // Load issues
    useEffect(() => {
        const load = async () => {
            await loadIssues(false);
        };
        load();
    }, []);

    const loadIssues = async (reload: boolean) => {
        try {
            setLoading(true);
            const { issueService, setting } = serviceContext;
            let allIssues: IssueCollection = null;
            for (const settingItem of setting.items) {
                const result = await issueService.getIssuesByQuery(settingItem, reload);

                if (result) {
                    if (allIssues == null) {
                        allIssues = cloneDeep(result);
                    } else {
                        //allIssues.items = allIssues.items.concat(result.items);
                        allIssues.appendIssues(result);
                    }
                    console.log(`settingItem: ${settingItem.id}; issues count: ${allIssues.items.length}`);
                }
            }
            if (allIssues && allIssues.items && allIssues.items.length > 0) {
                console.log("allIssues", allIssues);
                const groupResult = allIssues.createGroupsFromIssues();
                setGroups(groupResult);
                setIssues(allIssues);
            } else {
                setError("No issues was found.");
            }
        } catch (error) {
            setError(error);
        }
        setLoading(false);
    };

    const onRefreshHandler = async (): Promise<void> => {
        await loadIssues(true);
    }


    const onClickTaskHandler = (taskId: number): void => {
        props["routeProps"]["history"].push(`/taskitem/${taskId}`);
    }

    const _onIconClickHandler = async (task: Task): Promise<void> => {
        if (task.outlookMessage) {
            await task.outlookMessage.popupForm(isDialog);
        }
    }
    const theme = getTheme();
    const itemCell = {
        minHeight: 54,
        padding: 10,
        // boxSizing: 'border-box',
        borderBottom: `1px solid ${theme.semanticColors.bodyDivider}`,
        // display: 'flex',
        selectors: {
            '&:hover': { background: theme.palette.neutralLight },
        }
    };

    const containerStackTokens: IStackTokens = { childrenGap: 5 };

    const _onRenderCell = (nestingDepth?: number, item?: Issue, itemIndex?: number): React.ReactNode => {
        return item ? (
            <IssuesRow issue={item} />
        ) : null;
    };

    return (
        <div className="ms-Grid">
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 header" >
                    <h2>Issues</h2>
                </div>
            </div>
            <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <Stack tokens={containerStackTokens}>
                        <IssuesViewHeader
                            onRefresh={onRefreshHandler}
                        />
                        {loading ? <Spinner title="loading" />
                            :
                            <GroupedList items={issues && issues.items ? issues.items : []}
                                onRenderCell={_onRenderCell}
                                // groupProps={{ onRenderHeader: _onRenderHeader }}
                                selectionMode={SelectionMode.none}
                                compact={true}
                                groups={groups} />
                        }
                        {error && <div>{error}</div>}
                    </Stack>
                </div>
            </div>
        </div>
    );
};

// export default withAITracking(reactPlugin, IssuesView);
export default IssuesView;