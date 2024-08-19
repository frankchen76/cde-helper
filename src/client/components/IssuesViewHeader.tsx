import * as React from "react";
import { useState, useEffect, useContext } from "react";
import {
    ICommandBarItemProps,
    CommandBar
} from "@fluentui/react";
import { ServiceContext } from "../services/SettingService";

export interface IIssuesViewHeaderProps {
    onRefresh: () => void;
}
export const IssuesViewHeader = (props: IIssuesViewHeaderProps) => {
    const [cmdBarItems, setCmdbarItems] = useState<ICommandBarItemProps[]>([]);
    const serviceContext = useContext(ServiceContext);

    useEffect(() => {

        let cmdItemscmdBarItems: ICommandBarItemProps[] = [
            {
                key: 'refresh',
                text: 'Refresh',
                iconProps: { iconName: 'Refresh' },
                onClick: props.onRefresh
            }
        ];
        setCmdbarItems(cmdItemscmdBarItems);
    }, [serviceContext.setting]);

    return (
        <CommandBar
            styles={{ root: { padding: 0 } }}
            items={cmdBarItems}
            ariaLabel="Use left and right arrow keys to navigate between commands"
        />
    );
};