import { IStackTokens, Shimmer, ShimmerElementsGroup, ShimmerElementType, Stack } from "@fluentui/react";
import * as React from "react";
import { Common } from "../services/Common";

export const TaskItemShimmer = () => {
    const shimmerTitle = [
        { type: ShimmerElementType.line, width: "30%", height: 30 },
        { type: ShimmerElementType.gap, width: "70%", height: 30 }];
    const shimmerTextbox = [{ type: ShimmerElementType.line, width: "100%", height: 35 }];
    const shimmerTextarea = [{ type: ShimmerElementType.line, width: "100%", height: 160 }];
    const shimmerNumber = [
        { type: ShimmerElementType.line, width: "45%", height: 30 },
        { type: ShimmerElementType.gap, width: "10%", height: 30 },
        { type: ShimmerElementType.line, width: "45%", height: 30 }];
    const shimmerTogger = [
        { type: ShimmerElementType.line, width: "50%", height: 30 },
        { type: ShimmerElementType.gap, width: "20%", height: 30 },
        { type: ShimmerElementType.line, width: "30%", height: 30 }];
    const shimmerButtons = [
        { type: ShimmerElementType.gap, width: '12%', height: 30 },
        { type: ShimmerElementType.line, width: "25%", height: 30 },
        { type: ShimmerElementType.gap, width: '25%', height: 30 },
        { type: ShimmerElementType.line, width: "25%", height: 30 },
        { type: ShimmerElementType.gap, width: '12%', height: 30 }
    ];
    return (
        <Stack tokens={Common.CONTAINER_STACK_TOKENS}>
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextbox} />
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextbox} />
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextbox} />
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextbox} />
            <Shimmer shimmerElements={shimmerNumber} />
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextbox} />
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextarea} />
            <Shimmer shimmerElements={shimmerTogger} />
            <Shimmer shimmerElements={shimmerButtons} />
        </Stack>
    );
};
export const ShimmerCtrl = () => {
    return (
        <ShimmerElementsGroup
            flexWrap
            width="100%"
            shimmerElements={[
                { type: ShimmerElementType.line, width: "30%", height: 25 },
                { type: ShimmerElementType.gap, width: "70%", height: 25 },
                { type: ShimmerElementType.line, width: "100%", height: 30 }
            ]}
        />
    );
};
export const ShimmerButtons = () => {
    return (
        <ShimmerElementsGroup
            flexWrap
            width="100%"
            shimmerElements={[
                { type: ShimmerElementType.gap, width: "10%", height: 35 },
                { type: ShimmerElementType.line, width: "20%", height: 35 },
                { type: ShimmerElementType.gap, width: "10%", height: 35 },
                { type: ShimmerElementType.line, width: "20%", height: 35 },
                { type: ShimmerElementType.gap, width: "10%", height: 35 },
                { type: ShimmerElementType.line, width: "20%", height: 35 },
                { type: ShimmerElementType.gap, width: "10%", height: 35 }
            ]}
        />
    );
};
export const ShimmerQuill = () => {
    return (
        <div>
            <ShimmerElementsGroup
                flexWrap
                width="100%"
                shimmerElements={[
                    { type: ShimmerElementType.line, width: "30%", height: 25 },
                    { type: ShimmerElementType.gap, width: "70%", height: 25 },
                    { type: ShimmerElementType.gap, width: "100%", height: 25 }
                ]}
            />
            <ShimmerElementsGroup
                flexWrap
                width="100%"
                shimmerElements={[
                    { type: ShimmerElementType.line, width: "100%", height: 100 },
                    { type: ShimmerElementType.gap, width: "100%", height: 10 }
                ]}
            />
        </div>

    );
};
