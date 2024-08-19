import { CommandBar, ICommandBarItemProps } from "@fluentui/react";
import jwtDecode, { JwtPayload } from "jwt-decode";
import { find } from "lodash";
import * as React from "react";
import { useContext, useEffect, useState } from "react";
import { OutlookItem } from "../../services/OutlookItem";
import { ServiceContext } from "../../services/SettingService";
import { TaskCollection } from "../../services/Task";

export interface ITasksViewHeaderProps {
    outlookItem: OutlookItem;
    tasks: TaskCollection;
    isDialog: boolean;
}
export const TasksViewHeader = (props: ITasksViewHeaderProps) => {

    const [cmdBarItems, setCmdbarItems] = useState<ICommandBarItemProps[]>([]);
    const serviceContext = useContext(ServiceContext);


    useEffect(() => {
        const existingTask = props.tasks ? find(props.tasks.items, t => t.outlookMessage && t.outlookMessage.ItemId == props.outlookItem?.ItemId) : null;
        let btnText = props.outlookItem && props.outlookItem.ItemType == "message" ? "Add Email Task" : "Add Event Task";

        const subMailItems: ICommandBarItemProps[] = serviceContext.setting.items.map(settingItem => {
            return {
                key: `addemailtask-${settingItem.id}`,
                text: settingItem.name,
                iconProps: { iconName: 'NewMail' },
                href: `#/taskitem/${settingItem.id}/0`,
            };
        });
        const subTaskItems: ICommandBarItemProps[] = serviceContext.setting.items.map(settingItem => {
            return {
                key: `addemailtask-${settingItem.id}`,
                text: settingItem.name,
                iconProps: { iconName: 'PageAdd' },
                href: `#/taskitem/${settingItem.id}/-1`,
            };
        });
        let cmdItemscmdBarItems: ICommandBarItemProps[] = [
            {
                key: 'addemailtask',
                text: btnText,
                iconProps: { iconName: 'NewMail' },
                disabled: existingTask != null || props.isDialog,
                href: `#/taskitem/${serviceContext.setting.defaultSettingId}/0`,
                subMenuProps: { items: subMailItems }
            },
            {
                key: 'addtask',
                text: 'Add Task',
                iconProps: { iconName: 'PageAdd' },
                href: `#/taskitem/${serviceContext.setting.defaultSettingId}/-1`,
                subMenuProps: { items: subTaskItems }
            }//,
            // {
            //     key: 'test',
            //     text: 'Test',
            //     iconProps: { iconName: 'PageAdd' },
            //     onClick: onTestPara.bind(this, "test1") //onTestOutlookRESTHandler1
            // }
        ];
        setCmdbarItems(cmdItemscmdBarItems);
    }, [serviceContext.setting, props.outlookItem]);

    const onTestPara = (val: string): void => {
        console.log("val:", val);
    };
    const onTestCopyClipboard = (): void => {
        const elem = document.createElement('textarea')
        elem.value = "Hello world";

        document.body.append(elem)

        // Select the text and copy to clipboard
        elem.select()
        const success = document.execCommand('copy')
        elem.remove()

    };
    const onTestOutlookRESTHandler1 = (): void => {
        console.log("userProfile:");
        console.log(Office.context.mailbox.userProfile);
        Office.context.mailbox.getUserIdentityTokenAsync(function (result) {
            console.log("getUserIdentityTokenAsync:");
            console.log(result.value);
            const decodedToken = jwtDecode<JwtPayload>(result.value);
            //console.Console(`UPN: ${decodedToken.}`);
            console.log(decodedToken);


        });
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            console.log("getCallbackTokenAsync:");
            console.log(result.value);
            const decodedToken = jwtDecode<JwtPayload>(result.value);
            //console.Console(`UPN: ${decodedToken.}`);
            console.log(decodedToken);


        });
    };
    const onTestOutlookRESTHandler = (): void => {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            var ewsId = Office.context.mailbox.item.itemId;
            var token = result.value;
            var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
            var internetMessageId = Office.context.mailbox.item.internetMessageId;

            //var getMessageUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + restId;
            var getMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/inbox/messages?$filter=InternetMessageId eq '${internetMessageId}'`;

            console.log(getMessageUrl);
            console.log(token);

            var xhr = new XMLHttpRequest();
            xhr.open("GET", getMessageUrl);
            xhr.setRequestHeader("Authorization", "Bearer " + token);
            xhr.onload = function (e) {
                console.log(this.response);
                var jsonobj = JSON.parse(this.response);
                const itemId = jsonobj.value[0]["Id"];
                Office.context.mailbox.displayMessageForm(itemId);
            };
            xhr.send();
        });
    }
    const onTestHandler = (): void => {
        Office.context.mailbox.masterCategories.getAsync((result) => {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                console.log(result.value);
            }
        });
    }

    return (
        <CommandBar
            styles={{ root: { padding: 0 } }}
            items={cmdBarItems}
            ariaLabel="Use left and right arrow keys to navigate between commands"
        />
    );
};