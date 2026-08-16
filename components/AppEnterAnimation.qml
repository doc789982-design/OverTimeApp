import QtQuick

// ═══════════════════════════════════════════════════════════════════
// MD3 Emphasized Enter Animation
// Fade + Scale + Y Translation с cubic-bezier
// ═══════════════════════════════════════════════════════════════════
Item {
    id: animRoot
    
    width: 0
    height: 0
    visible: false

    property Item target: null
    property int delay: 0

    // MD3 Emphasized Decelerate: 350ms
    property int duration: 350
    
    // Начальные значения
    property real scaleFrom: 0.92
    property real yOffsetFrom: 8

    // MD3 Emphasized Decelerate bezier
    // cubic-bezier(0.05, 0, 0.2, 1)
    readonly property var md3Decelerate: [0.05, 0, 0.2, 1, 1, 1]

    function enter() {
        if (!target) return
        staggerTimer.restart()
    }

    function reset() {
        if (!target) return
        staggerTimer.stop()
        enterAnim.stop()
        target.opacity = 0.0
        target.scale = animRoot.scaleFrom
        target.y = target.y + animRoot.yOffsetFrom
    }

    function resetAndEnter() {
        reset()
        staggerTimer.interval = animRoot.delay
        staggerTimer.restart()
    }

    Timer {
        id: staggerTimer
        interval: animRoot.delay
        repeat: false
        onTriggered: enterAnim.start()
    }

    ParallelAnimation {
        id: enterAnim

        // FADE
        NumberAnimation {
            target: animRoot.target
            property: "opacity"
            from: 0.0
            to: 1.0
            duration: animRoot.duration
            easing.type: Easing.BezierSpline
            easing.bezierCurve: animRoot.md3Decelerate
        }

        // SCALE
        NumberAnimation {
            target: animRoot.target
            property: "scale"
            from: animRoot.scaleFrom
            to: 1.0
            duration: animRoot.duration
            easing.type: Easing.BezierSpline
            easing.bezierCurve: animRoot.md3Decelerate
        }

        // Y TRANSLATION (slide up)
        NumberAnimation {
            target: animRoot.target
            property: "y"
            from: target.y + animRoot.yOffsetFrom
            to: target.y
            duration: animRoot.duration
            easing.type: Easing.BezierSpline
            easing.bezierCurve: animRoot.md3Decelerate
        }
    }
}
