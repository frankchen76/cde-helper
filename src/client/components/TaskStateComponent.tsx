import * as React from "react";
import { useState, useEffect } from "react";

export interface TaskStateComponentProps {
    state: string;
}

export const TaskStateComponent = (props: TaskStateComponentProps) => {
    const barStyle = { width: 6, fontSize: 15 };
    const [bars, setBars] = useState<string[]>();
    useEffect(() => {
        switch (props.state.toLowerCase()) {
            case "to do":
                setBars(["ms-Icon ms-Icon--WorkItemBar", "ms-Icon ms-Icon--WorkItemBar", "ms-Icon ms-Icon--WorkItemBar"]);
                break;
            case "doing":
                setBars(["ms-Icon ms-Icon--WorkItemBarSolid", "ms-Icon ms-Icon--WorkItemBarSolid", "ms-Icon ms-Icon--WorkItemBar"]);
                break;
            case "done":
                setBars(["ms-Icon ms-Icon--WorkItemBarSolid", "ms-Icon ms-Icon--WorkItemBarSolid", "ms-Icon ms-Icon--WorkItemBarSolid"]);
                break;
        }
    }, [props.state]);
    return (
        <div title={props.state}>
            {bars && bars.map((b, index) => <span className={b} key={`taskStateComponent_${index}`} style={barStyle} aria-hidden="true"></span>)}
        </div>
    )
};