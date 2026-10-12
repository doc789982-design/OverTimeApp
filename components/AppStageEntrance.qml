import QtQuick

// ═══════════════════════════════════════════════════════════════════
// УВЕРТЮРА: один участник выхода на сцену.
//
// Прикрепляется к любому элементу и в свой ход (delay) проявляет его
// с лёгким смещением по своей оси: сетка месяца поднимается снизу,
// левая панель въезжает слева. Смещение делается через Translate-
// трансформацию — раскладка (SplitView, Column, StackLayout) не
// затрагивается, ничего не перестраивается.
//
// Канон: вход = торможение (easeEnter), 250 мс.
// ═══════════════════════════════════════════════════════════════════
Item {
    id: root

    width: 0
    height: 0
    visible: false

    // Кому выходить на сцену. ВНИМАНИЕ: play() назначает цели свою
    // трансформацию — у цели не должно быть своих transform-ов.
    property Item target: null

    // Ход участника: задержка от начала увертюры (мс)
    property int delay: 0

    // Откуда входит: слева/справа (fromX) и/или снизу (fromY)
    property real fromX: 0
    property real fromY: AppTheme.riseOffset

    property int duration: AppTheme.durStandard

    readonly property Translate shift: Translate { id: shift }

    function play() {
        if (!target) return
        safety.stop()
        target.transform = [shift]
        target.opacity = 0
        shift.x = fromX
        shift.y = fromY
        enterAnim.restart()
        safety.restart()
    }

    // Страховка: даже если анимация не отбежала — элемент всё равно
    // оказывается на месте и виден (как и в старом «подъезде»)
    function forceFinish() {
        enterAnim.stop()
        shift.x = 0
        shift.y = 0
        if (target) target.opacity = 1
    }

    SequentialAnimation {
        id: enterAnim
        PauseAnimation { duration: root.delay }
        ParallelAnimation {
            NumberAnimation {
                target: root.target; property: "opacity"
                to: 1; duration: root.duration
                easing.type: AppTheme.easeEnter
            }
            NumberAnimation {
                target: shift; property: "x"
                to: 0; duration: root.duration
                easing.type: AppTheme.easeEnter
            }
            NumberAnimation {
                target: shift; property: "y"
                to: 0; duration: root.duration
                easing.type: AppTheme.easeEnter
            }
        }
    }

    Timer {
        id: safety
        interval: root.delay + root.duration + 150
        onTriggered: root.forceFinish()
    }
}
