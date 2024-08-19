import { IStackTokens, Shimmer, ShimmerElementType, Stack } from "@fluentui/react";
import * as React from "react";
import { Common } from "../../services/Common";

export const IssueItemShimmer = () => {
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
            <Shimmer shimmerElements={shimmerTitle} />
            <Shimmer shimmerElements={shimmerTextarea} />
            <Shimmer shimmerElements={shimmerButtons} />
        </Stack>
    );
};
