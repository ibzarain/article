"use strict";
Object.defineProperty(exports, "__esModule", {
    value: true
});
Object.defineProperty(exports, "rotateAtom", {
    enumerable: true,
    get: function() {
        return rotateAtom;
    }
});
const _reactmotion = require("@fluentui/react-motion");
const createRotateValue = (axis, angle)=>{
    return `${axis.toLowerCase()} ${angle}deg`;
};
const rotateAtom = ({ direction, duration, easing = _reactmotion.motionTokens.curveLinear, delay = 0, axis = 'z', fromAngle = -90, toAngle = 0 })=>{
    const keyframes = [
        {
            rotate: createRotateValue(axis, fromAngle)
        },
        {
            rotate: createRotateValue(axis, toAngle)
        }
    ];
    if (direction === 'exit') {
        keyframes.reverse();
    }
    return {
        keyframes,
        duration,
        easing,
        delay
    };
};
