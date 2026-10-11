import QtQuick
import "."

// ============================================================
// СЦЕНА-ТЕАТР ДЛЯ ЖИВОЙ ВИНЬЕТКИ «ЧТО НОВОГО»
//
// Декорации и актёры маленького спектакля про изменение выпуска:
// дни календаря с метками, курсор, меню статусов, мини-окно —
// и «плеер сценариев». Сцена не собирается вручную: она записывается
// списком тактов «что за чем происходит», а плеер проигрывает их
// один за другим, по кругу. Новая сцена — это новый список тактов
// из существующих глаголов (примитивов), без новой анимационной
// логики.
//
// Все фигурки — в цветах темы (подхватывается автоматически),
// те же кривые и длительности, что у остального интерфейса.
// ============================================================
Item {
    id: stage
    objectName: "whatsNewStage"

    width: parent ? parent.width : 0
    height: parent ? parent.height : 0

    // Какую букву ставят метки в этой сцене («О» — отпуск)
    property string badgeLetter: "О"

    // ══════════════════════════════════════════════════════════
    // ДЕКОРАЦИИ: семь дней, как неделя в календаре
    // ══════════════════════════════════════════════════════════
    readonly property int dayCount: 7
    readonly property real cellW: 48
    readonly property real cellH: 64
    readonly property real cellGap: 8
    readonly property real stripW: dayCount * cellW + (dayCount - 1) * cellGap
    readonly property real rowY: (height - cellH) / 2

    function dayX(i) { return (width - stripW) / 2 + i * (cellW + cellGap) }
    function dayCenterX(i) { return dayX(i) + cellW / 2 }

    // Цвета метки по букве — как у меток календаря
    function badgeColors(letter) {
        if (letter === "Б") return { "bg": AppTheme.bgDangerSoft, "fg": AppTheme.accentDanger }
        if (letter === "К") return { "bg": AppTheme.bgPurpleSoft, "fg": AppTheme.accentPurple }
        if (letter === "В") return { "bg": AppTheme.bgTealSoft, "fg": AppTheme.accentTeal }
        return { "bg": AppTheme.bgWarningSoft, "fg": AppTheme.accentWarning }
    }

    Row {
        x: (stage.width - stage.stripW) / 2
        y: stage.rowY
        spacing: stage.cellGap

        Repeater {
            id: daysRepeater
            model: stage.dayCount

            Rectangle {
                id: miniCell
                objectName: "vignetteMiniCell"
                width: stage.cellW
                height: stage.cellH
                radius: AppTheme.radiusMedium
                // выходные (числа 10 и 11) — как в календаре, красноватые
                color: (5 + index) >= 10
                       ? AppTheme.bgDangerSoft : AppTheme.bgCell

                // фасады для сценария
                function resetBadge() { miniBadge.resetBadge() }
                function startBadgeIn() { miniBadge.startIn() }
                function startBadgeOut() { miniBadge.startOut() }

                // число дня — маленькое, в правом верхнем углу
                Text {
                    anchors.top: parent.top
                    anchors.right: parent.right
                    anchors.margins: 5
                    text: 5 + index
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    font.weight: AppTheme.weightBold
                    color: (5 + index) >= 10
                           ? AppTheme.accentDanger
                           : AppTheme.textSecondary
                }

                // плёнка выделения — следует за курсором
                Rectangle {
                    anchors.fill: parent
                    radius: parent.radius
                    color: AppTheme.stateHover
                    opacity: (cursor.sweeping
                              && cursor.x + cursor.width / 2
                                 >= stage.dayCenterX(index)) ? 1 : 0
                    Behavior on opacity {
                        NumberAnimation { duration: AppTheme.durMicro }
                    }
                }

                // метка статуса — под числом, как в ячейке
                Rectangle {
                    id: miniBadge
                    objectName: "vignetteMiniBadge"
                    anchors.right: parent.right
                    anchors.rightMargin: 5
                    y: 26
                    width: 16
                    height: 16
                    radius: AppTheme.radiusPill
                    color: stage.badgeColors(stage.badgeLetter).bg
                    opacity: 0
                    scale: 0.55

                    Text {
                        anchors.centerIn: parent
                        text: stage.badgeLetter
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeMicro
                        font.weight: AppTheme.weightBold
                        color: stage.badgeColors(stage.badgeLetter).fg
                    }

                    function resetBadge() {
                        stopBadgeAnims()
                        opacity = 0
                        scale = 0.55
                    }
                    function stopBadgeAnims() {
                        badgeIn.stop()
                        badgeOut.stop()
                    }
                    // старт из чужого скоупа: id-шники анимаций —
                    // не свойства, зовём только методы
                    function startIn() {
                        stopBadgeAnims()
                        badgeIn.restart()
                    }
                    function startOut() {
                        stopBadgeAnims()
                        badgeOut.restart()
                    }

                    // сборка метки: проявление и рост, очередь волны —
                    // задержка по порядку дня
                    SequentialAnimation {
                        id: badgeIn
                        PauseAnimation { duration: 80 + index * 40 }
                        ParallelAnimation {
                            NumberAnimation { target: miniBadge; property: "opacity"; to: 1; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
                            NumberAnimation { target: miniBadge; property: "scale"; to: 1; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
                        }
                    }
                    // растворение: выход короче входа
                    ParallelAnimation {
                        id: badgeOut
                        NumberAnimation { target: miniBadge; property: "opacity"; to: 0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
                        NumberAnimation { target: miniBadge; property: "scale"; to: 0.55; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
                    }
                }
            }
        }
    }

    // ══════════════════════════════════════════════════════════
    // АКТЁР: курсор — точка, которую «ведут» по сцене
    // ══════════════════════════════════════════════════════════
    Rectangle {
        id: cursor
        objectName: "vignetteCursor"
        property bool sweeping: false
        property real pressScale: 1.0
        width: 12
        height: 12
        radius: 6
        x: stage.dayCenterX(0) - width / 2
        y: stage.rowY + stage.cellH / 2 - height / 2
        color: AppTheme.accentBrand
        opacity: 0
        scale: pressScale
        Behavior on opacity {
            NumberAnimation { duration: AppTheme.durMicro }
        }
    }
    NumberAnimation { id: cursorMove; target: cursor; property: "x"; duration: 800; easing.type: Easing.InOutSine }
    NumberAnimation { id: cursorMoveY; target: cursor; property: "y"; duration: 300; easing.type: Easing.InOutSine }
    SequentialAnimation {
        id: cursorPressAnim
        NumberAnimation { target: cursor; property: "pressScale"; to: 0.75; duration: 80; easing.type: Easing.OutQuad }
        NumberAnimation { target: cursor; property: "pressScale"; to: 1.0; duration: 120; easing.type: Easing.OutQuad }
    }

    // ══════════════════════════════════════════════════════════
    // АКТЁР: меню статусов — мини-копия меню дня
    // (шапка с диапазоном дат и кружки К/Б/О + корзина)
    // ══════════════════════════════════════════════════════════
    Rectangle {
        id: menu
        objectName: "vignetteMenu"
        z: 5
        visible: opacity > 0.01
        opacity: 0
        scale: 0.96
        width: 156
        height: 74
        radius: AppTheme.radiusMedium
        color: AppTheme.bgElevated
        border.color: AppTheme.borderDivider
        border.width: 1
        property int hoverIdx: -1
        property int pressIdx: -1
        Behavior on opacity {
            NumberAnimation { duration: AppTheme.durMicro }
        }
        Behavior on scale {
            NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeEnter }
        }

        // шапка: диапазон выделенных дней + закрыть
        Text {
            anchors.left: parent.left
            anchors.leftMargin: 10
            anchors.top: parent.top
            anchors.topMargin: 6
            text: "05–11.10"
            font.family: AppTheme.fontCondensed
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightBold
            color: AppTheme.textPrimary
        }
        Rectangle {
            width: 16; height: 16; radius: 8
            anchors.right: parent.right
            anchors.rightMargin: 8
            anchors.top: parent.top
            anchors.topMargin: 5
            color: AppTheme.stateHover
            Text {
                anchors.centerIn: parent
                text: "×"
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightBold
                color: AppTheme.textSecondary
            }
        }

        // кружки статусов
        Row {
            id: menuCircles
            anchors.top: parent.top
            anchors.topMargin: 30
            anchors.horizontalCenter: parent.horizontalCenter
            spacing: 8

            Repeater {
                model: [
                    { "t": "К", "fg": AppTheme.accentPurple },
                    { "t": "Б", "fg": AppTheme.accentDanger },
                    { "t": "О", "fg": AppTheme.accentWarning },
                    { "t": "", "fg": AppTheme.textSecondary }
                ]

                Rectangle {
                    width: 24
                    height: 24
                    radius: 12
                    color: menu.pressIdx === index ? AppTheme.statePress
                          : (menu.hoverIdx === index ? AppTheme.stateHover : AppTheme.bgSurface)
                    border.color: menu.hoverIdx === index
                                  ? AppTheme.accentBrand : AppTheme.borderDivider
                    border.width: 1
                    scale: menu.pressIdx === index ? 0.92 : 1.0
                    Behavior on scale {
                        NumberAnimation { duration: AppTheme.durMicro }
                    }

                    Text {
                        anchors.centerIn: parent
                        visible: modelData.t !== ""
                        text: modelData.t
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        font.weight: AppTheme.weightBold
                        color: modelData.fg
                    }
                    IconImage {
                        anchors.centerIn: parent
                        visible: modelData.t === ""
                        source: "../icons/trash.svg"
                        width: 12
                        height: 12
                        color: modelData.fg
                    }
                }
            }
        }
    }

    // ══════════════════════════════════════════════════════════
    // АКТЁР: мини-окно (для сцен «открыл окно, нажал кнопку»)
    // ══════════════════════════════════════════════════════════
    Rectangle {
        id: miniWindow
        objectName: "vignetteWindow"
        z: 6
        visible: opacity > 0.01
        opacity: 0
        scale: 0.92
        width: 176
        height: 92
        x: (stage.width - width) / 2
        y: (stage.height - height) / 2
        radius: AppTheme.radiusMedium
        color: AppTheme.bgModal
        border.color: AppTheme.borderDivider
        border.width: 1
        Behavior on opacity {
            NumberAnimation { duration: AppTheme.durMicro }
        }
        Behavior on scale {
            NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeEnter }
        }

        Text {
            anchors.left: parent.left
            anchors.leftMargin: 12
            anchors.top: parent.top
            anchors.topMargin: 8
            text: "Окно"
            font.family: AppTheme.fontCondensed
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightBold
            color: AppTheme.textPrimary
        }
        Rectangle {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: parent.top
            anchors.topMargin: 26
            height: 1
            color: AppTheme.borderDivider
        }

        Row {
            anchors.bottom: parent.bottom
            anchors.bottomMargin: 10
            anchors.horizontalCenter: parent.horizontalCenter
            spacing: 8

            Rectangle {
                width: 64
                height: 22
                radius: AppTheme.radiusSmall
                color: AppTheme.bgSurface
                border.color: AppTheme.borderDivider
                border.width: 1
                Text {
                    anchors.centerIn: parent
                    text: "Отмена"
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeMicro
                    color: AppTheme.textPrimary
                }
            }
            Rectangle {
                id: winSave
                objectName: "vignetteSaveButton"
                property real pressScale: 1.0
                width: 76
                height: 22
                radius: AppTheme.radiusSmall
                color: AppTheme.accentBrand
                scale: pressScale
                Text {
                    anchors.centerIn: parent
                    text: "Сохранить"
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeMicro
                    font.weight: AppTheme.weightBold
                    color: "#FFFFFF"
                }
            }
        }
    }
    SequentialAnimation {
        id: savePressAnim
        NumberAnimation { target: winSave; property: "pressScale"; to: 0.94; duration: 80; easing.type: Easing.OutQuad }
        NumberAnimation { target: winSave; property: "pressScale"; to: 1.0; duration: 120; easing.type: Easing.OutQuad }
    }

    // ══════════════════════════════════════════════════════════
    // ПЛЕЕР СЦЕНАРИЕВ: такт за тактом, по кругу
    // ══════════════════════════════════════════════════════════
    property var steps: []
    property int stepIndex: 0
    property bool playing: false

    // Полная длительность прогона — карусель по ней листает сцены
    readonly property int totalMs: {
        let t = 0
        for (let i = 0; i < steps.length; i++)
            t += (steps[i] && steps[i].ms) ? steps[i].ms : 0
        return t
    }

    Timer {
        id: stepTimer
        interval: 10
        onTriggered: { stage.stepIndex++; stage.runStep() }
    }

    function playFromStart() {
        resetAll()
        stepIndex = 0
        playing = true
        runStep()
    }
    function stopPlay() {
        playing = false
        stepTimer.stop()
    }
    function runStep() {
        if (!playing || steps.length === 0) return
        if (stepIndex >= steps.length) stepIndex = 0     // сюжет по кругу
        let s = steps[stepIndex] || {}
        if (s.run) s.run()
        stepTimer.interval = Math.max(1, s.ms || 1)
        stepTimer.restart()
    }

    // ══════════════════════════════════════════════════════════
    // ГЛАГОЛЫ СЦЕНЫ (примитивы, из которых записывают сценарий)
    // ══════════════════════════════════════════════════════════
    function resetAll() {
        cursor.sweeping = false
        cursor.opacity = 0
        cursor.x = dayCenterX(0) - cursor.width / 2
        cursor.y = rowY + cellH / 2 - cursor.height / 2
        cursor.pressScale = 1
        for (let i = 0; i < daysRepeater.count; i++) {
            let d = daysRepeater.itemAt(i)
            if (d) d.resetBadge()
        }
        menuClose()
        windowClose()
        winSave.pressScale = 1
    }

    function cursorShow() { cursor.opacity = 0.9 }
    function cursorHide() { cursor.opacity = 0; cursor.sweeping = false }
    function cursorSweepTo(dayIdx) {
        cursor.sweeping = true
        cursorMove.duration = 800
        cursorMove.to = dayCenterX(dayIdx) - cursor.width / 2
        cursorMove.restart()
    }
    function cursorTo(x, y, ms) {
        cursorMove.to = x
        cursorMove.duration = ms
        cursorMove.restart()
        cursorMoveY.to = y
        cursorMoveY.duration = ms
        cursorMoveY.restart()
    }
    function cursorPress() { cursorPressAnim.restart() }

    function menuCircleCenter(i) {
        let localX = menuCircles.x + i * (24 + menuCircles.spacing) + 12
        let pt = menu.mapToItem(stage, localX, menuCircles.y + 12)
        return pt
    }
    function cursorToMenu(i) {
        let pt = menuCircleCenter(i)
        cursorTo(pt.x - cursor.width / 2, pt.y - cursor.height / 2, 300)
    }
    function menuOpen(dayIdx) {
        menu.x = Math.max(4, Math.min(dayCenterX(dayIdx) - menu.width / 2,
                                      width - menu.width - 4))
        menu.y = Math.max(4, Math.min(rowY + 8, height - menu.height - 4))
        menu.opacity = 1
        menu.scale = 1
    }
    function menuHighlight(i) { menu.hoverIdx = i }
    function menuPress(i) { menu.pressIdx = i }
    function menuClose() {
        menu.opacity = 0
        menu.scale = 0.96
        menu.hoverIdx = -1
        menu.pressIdx = -1
    }

    function badgesWave() {
        for (let i = 0; i < daysRepeater.count; i++) {
            let d = daysRepeater.itemAt(i)
            if (d) d.startBadgeIn()
        }
    }
    function badgesDissolve() {
        for (let i = 0; i < daysRepeater.count; i++) {
            let d = daysRepeater.itemAt(i)
            if (d) d.startBadgeOut()
        }
    }
    function clearHighlights() { cursor.sweeping = false }

    function windowOpen() { miniWindow.opacity = 1; miniWindow.scale = 1 }
    function windowClose() { miniWindow.opacity = 0; miniWindow.scale = 0.92 }
    function windowPressSave() { savePressAnim.restart() }
}
