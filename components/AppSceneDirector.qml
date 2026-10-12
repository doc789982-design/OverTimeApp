import QtQuick

// ═══════════════════════════════════════════════════════════════════
// РЕЖИССЁР СЦЕН (общая ось, Material 3 shared axis).
//
// Смена содержимого внутри одного окна или панели: старая сцена
// уходит с разгоном (короткий выход, easeExit), новая входит с
// торможением по той же оси (вход, easeEnter). Смена значения
// происходит ТОЛЬКО через swap(): пока идёт выход, подмена отложена,
// поэтому старое содержание честно уходит на глазах.
//
// Сцены: разделы настроек, год производственного календаря,
// месяц ↔ год в календаре. Смещение — Translate-трансформацией,
// раскладка не затрагивается.
// ═══════════════════════════════════════════════════════════════════
Item {
    id: root

    width: 0
    height: 0
    visible: false

    // Контейнер сцены (меняется его содержимое). ВНИМАНИЕ: swap()
    // назначает контейнеру свою трансформацию — своих transform-ов
    // у него быть не должно.
    property Item target: null

    // Сцены играются только на экране; закрытое окно меняет
    // содержимое мгновенно, без хореографии вхолостую
    // (armed, а не enabled: enabled уже есть у Item)
    property bool armed: true

    // Общая ось сцены: уход −axisY, вход +axisY (0 = только проявление)
    property real axisX: 0
    property real axisY: AppTheme.riseOffset

    // Сдвиги конкретного обмена (выставляются в swap)
    property real _exitShiftX: 0
    property real _exitShiftY: 0
    property real _enterShiftX: 0
    property real _enterShiftY: 0

    // Отложенная подмена содержимого
    property var pending: null

    readonly property Translate shift: Translate { id: shift }

    // Главное окно программы меняет сцену только через swap().
    // mutator — функция, подменяющая содержимое (смена индекса,
    // года и т.п.). exitShift/enterShift — свои сдвиги для этого
    // обмена (по умолчанию — ось сцены).
    function swap(mutator, exitShiftX, exitShiftY, enterShiftX, enterShiftY) {
        _exitShiftX = (exitShiftX === undefined) ? axisX : exitShiftX
        _exitShiftY = (exitShiftY === undefined) ? axisY : exitShiftY
        _enterShiftX = (enterShiftX === undefined) ? axisX : enterShiftX
        _enterShiftY = (enterShiftY === undefined) ? axisY : enterShiftY

        if (!target || !armed) {
            if (mutator) mutator()
            return
        }

        // Прежний обмен ещё не доигран — завершаем мгновенно
        // (его подмена не состоялась и отменяется)
        pending = null
        exitAnim.stop()
        enterAnim.stop()

        pending = mutator
        target.transform = [shift]
        exitAnim.restart()
        safety.restart()
    }

    // Страховка: подмена всё равно применяется, сцена оказывается
    // на месте и видна
    function forceFinish() {
        if (pending) {
            let m = pending
            pending = null
            m()
        }
        exitAnim.stop()
        enterAnim.stop()
        shift.x = 0
        shift.y = 0
        if (target) target.opacity = 1
    }

    SequentialAnimation {
        id: exitAnim
        ParallelAnimation {
            NumberAnimation {
                target: root.target; property: "opacity"
                to: 0; duration: AppTheme.durFast
                easing.type: AppTheme.easeExit
            }
            NumberAnimation {
                target: shift; property: "x"
                to: -root._exitShiftX; duration: AppTheme.durFast
                easing.type: AppTheme.easeExit
            }
            NumberAnimation {
                target: shift; property: "y"
                to: -root._exitShiftY; duration: AppTheme.durFast
                easing.type: AppTheme.easeExit
            }
        }
        onStopped: {
            if (pending) {
                let m = pending
                pending = null
                m()
                enterAnim.restart()
            }
        }
    }

    SequentialAnimation {
        id: enterAnim
        ParallelAnimation {
            NumberAnimation {
                target: root.target; property: "opacity"
                to: 1; duration: AppTheme.durStandard
                easing.type: AppTheme.easeEnter
            }
            NumberAnimation {
                target: shift; property: "x"
                from: root._enterShiftX; to: 0; duration: AppTheme.durStandard
                easing.type: AppTheme.easeEnter
            }
            NumberAnimation {
                target: shift; property: "y"
                from: root._enterShiftY; to: 0; duration: AppTheme.durStandard
                easing.type: AppTheme.easeEnter
            }
        }
    }

    Timer {
        id: safety
        interval: AppTheme.durFast + AppTheme.durStandard + 200
        onTriggered: root.forceFinish()
    }
}
