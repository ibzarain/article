"use strict";
Object.defineProperty(exports, "__esModule", {
    value: true
});
Object.defineProperty(exports, "slideAtom", {
    enumerable: true,
    get: function() {
        return slideAtom;
    }
});
const _reactmotion = require("@fluentui/react-motion");
const slideAtom = ({ direction, duration, easing = _reactmotion.motionTokens.curveLinear, delay = 0, fromX = '0px', fromY = '0px', toX = '0px', toY = '0px' })=>{
    const keyframes = [
        {
            translate: `${fromX} ${fromY}`
        },
        {
            translate: `${toX} ${toY}`
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
