import QtQuick
import QtQuick.Controls.impl

// ============================================================
// ОКНО ДНЯ (открывается ЛКМ по ячейке календаря)
//
// Появляется как контекстное меню: рамка с тенью и плавное появление
// (прозрачность + масштаб), позиция — у ячейки дня. Клик по пункту —
// окно дорастает до следующего окна (дежурство/компенсация/«Открыть день»).
// При «ОК» / «Отмена» окно закрывается целиком — старые не выскакивают.
// ============================================================
Item {
    id: root

    anchors.fill: parent
    z: AppTheme.zModal - 1

    // Видимость: окно скрыто (0), меню (1), растёт в диалог (2)
    property int mode: 0
    property string targetDate: ""

    property bool menuIsWeekend: false
    property bool menuIsHoliday: false
    property bool menuHasDuties: false
    property bool menuHasComps: false

    property real menuW: 256

    // Геометрия ячейки, из которой открылись (для обратной анимации)
    property rect cellRect: Qt.rect(0, 0, 0, 0)

    // Длительности — как в первой итерации (плавно, но не дольше)
    readonly property int durGrow: AppTheme.durSlow
    readonly property int durContent: AppTheme.durStandard

    visible: boxOpacity > 0
    opacity: 1

    // Фон всего экрана: клик мимо закрывает меню (только в режиме меню)
    MouseArea {
        anchors.fill: parent
        enabled: root.mode === 1
        onClicked: root.closeFromCell()
    }

    // ============================================================
    // САМО ОКНО (анимируемая рамка)
    // ============================================================
    Item {
        id: box
        x: 0; y: 0
        width: 0; height: 0
        opacity: root.boxOpacity
        scale: 1.0
        transformOrigin: Item.TopLeft
        Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }

        Rectangle {
            anchors.fill: parent
            color: AppTheme.bgModal
            radius: AppTheme.radiusModal
            border.color: AppTheme.borderDivider
            border.width: 1
            // Тень-картинка вместо вычисляемой
            AppShadow { level: 4 }
        }

        // ============================================================
        // СОДЕРЖИМОЕ МЕНЮ (клипится, чтобы во время роста не вылезать
        // за рамку, а тень осталась видна)
        // ============================================================
        Item {
            id: contentClip
            anchors.fill: parent
            clip: true

            Column {
                id: menuContent
                x: AppTheme.spaceM
                y: AppTheme.spaceM
                width: root.menuW - 2 * AppTheme.spaceM
                spacing: 4
                opacity: 0
                scale: 0.97
                transformOrigin: Item.TopLeft
                Behavior on opacity { NumberAnimation { duration: root.durContent; easing.type: AppTheme.easeStandard } }
                Behavior on scale { NumberAnimation { duration: root.durContent; easing.type: AppTheme.easeStandard } }

                // ── Шапка: дата + закрыть ──
                Row {
                    width: parent.width
                    spacing: AppTheme.spaceS
                    Text {
                        id: dateLabel
                        width: parent.width - 40
                        elide: Text.ElideRight
                        text: root.fmtDate(root.targetDate)
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontCondensed
                        font.pixelSize: AppTheme.sizeH4
                        font.weight: AppTheme.weightBold
                        verticalAlignment: Text.AlignVCenter
                        height: 34
                    }
                    Rectangle {
                        width: 28; height: 28; radius: AppTheme.radiusPill
                        anchors.verticalCenter: parent.verticalCenter
                        color: closeHov.pressed ? AppTheme.statePress
                              : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                        IconImage { anchors.centerIn: parent; source: "../icons/close.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.textSecondary }
                        MouseArea { id: closeHov; anchors.fill: parent; hoverEnabled: true; onClicked: root.closeFromCell() }
                    }
                }

                // ── Быстрые статусы: К / Б / О / сброс (круглые) ──
                Row {
                    width: parent.width
                    spacing: AppTheme.spaceS
                    Repeater {
                        model: [
                            { "t": "К", "tool": "Командировка", "c": AppTheme.accentPurple, "st": "К" },
                            { "t": "Б", "tool": "Больничный",     "c": AppTheme.accentDanger, "st": "Б" },
                            { "t": "О", "tool": "Отпуск",         "c": AppTheme.accentWarning, "st": "О" }
                        ]
                        Rectangle {
                            width: 34; height: 34; radius: AppTheme.radiusPill
                            color: m.pressed ? AppTheme.statePress
                                  : (m.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
                            border.color: AppTheme.borderDivider
                            border.width: 1
                            Text {
                                anchors.centerIn: parent
                                text: modelData.t
                                color: modelData.c
                                font.weight: AppTheme.weightBold
                                font.pixelSize: AppTheme.sizeBodyLarge
                            }
                            MouseArea {
                                id: m
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onClicked: {
                                    backend.setDayStatus(root.targetDate, modelData.st)
                                    root.closeFromCell()
                                }
                            }
                            AppToolTip {
                                anchors.horizontalCenter: parent.horizontalCenter
                                anchors.bottom: parent.top
                                anchors.bottomMargin: AppTheme.spaceXS
                                text: modelData.tool
                                isVisible: m.containsMouse
                            }
                        }
                    }

                    Rectangle {
                        width: 34; height: 34; radius: AppTheme.radiusPill
                        color: clearBtn.pressed ? AppTheme.statePress
                              : (clearBtn.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
                        border.color: AppTheme.borderDivider
                        border.width: 1
                        IconImage {
                            anchors.centerIn: parent
                            source: "../icons/trash.svg"
                            width: AppTheme.iconMedium
                            height: AppTheme.iconMedium
                            color: clearBtn.containsMouse ? AppTheme.accentDanger : AppTheme.textSecondary
                        }
                        MouseArea {
                            id: clearBtn
                            anchors.fill: parent
                            hoverEnabled: true
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                mainWindow.explodeAndDelete(root.targetDate, "status", null,
                                    function() { backend.setDayStatus(root.targetDate, "") })
                                root.closeFromCell()
                            }
                        }
                        AppToolTip {
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.bottom: parent.top
                            anchors.bottomMargin: AppTheme.spaceXS
                            text: "Удалить статус"
                            isVisible: clearBtn.containsMouse
                        }
                    }
                }

                Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

                // ── Пункты меню ──
                AppMenuRow {
                    iconSource: "../icons/edit.svg"
                    text: "Открыть день"
                    onClicked: {
                        backend.loadDayDetails(root.targetDate)
                        dayInspector.targetDate = root.targetDate
                        root.morphToDialog(450, dayInspector.effectiveHeight,
                            function(x, y, w, h) { dayInspector.openMorph(x, y, w, h) })
                    }
                }

                AppMenuRow {
                    iconSource: "../icons/clock.svg"
                    text: "Добавить дежурство"
                    showDelete: root.menuHasDuties
                    onClicked: {
                        dayDutyDialog.prepareForDuty(root.targetDate)
                        root.morphToDialog(380, dayDutyDialog.effectiveHeight,
                            function(x, y, w, h) { dayDutyDialog.openForDutyMorph(root.targetDate, x, y, w, h) })
                    }
                    onDeleteClicked: {
                        mainWindow.askConfirm(
                            "Удалить все дежурства?",
                            "Будут удалены все дежурства за " + root.fmtDate(root.targetDate) + ".\nЕсли передумаете — нажмите Ctrl+Z.",
                            "Удалить",
                            function() {
                                mainWindow.explodeAndDelete(root.targetDate, "duty", null,
                                    function() { backend.clearDayDuties(root.targetDate) })
                            }
                        )
                        root.closeFromCell()
                    }
                }

                AppMenuRow {
                    iconSource: "../icons/rest.svg"
                    text: "Добавить компенсацию"
                    showDelete: root.menuHasComps
                    onClicked: {
                        dayCompDialog.prepareForComp(root.targetDate)
                        root.morphToDialog(380, dayCompDialog.effectiveHeight,
                            function(x, y, w, h) { dayCompDialog.openForCompMorph(root.targetDate, x, y, w, h) })
                    }
                    onDeleteClicked: {
                        mainWindow.askConfirm(
                            "Удалить все компенсации?",
                            "Будут удалены все компенсации за " + root.fmtDate(root.targetDate) + ".\nЕсли передумаете — нажмите Ctrl+Z.",
                            "Удалить",
                            function() {
                                mainWindow.explodeAndDelete(root.targetDate, "comp", null,
                                    function() { backend.clearDayCompensations(root.targetDate) })
                            }
                        )
                        root.closeFromCell()
                    }
                }

                Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

                AppMenuRow {
                    visible: root.menuIsWeekend || root.menuIsHoliday
                    iconSource: "../icons/calendar.svg"
                    text: "Сделать рабочим"
                    customColor: AppTheme.accentTeal
                    onClicked: {
                        backend.setDayType(root.targetDate, "work")
                        root.closeFromCell()
                    }
                }
                AppMenuRow {
                    visible: !root.menuIsWeekend
                    iconSource: "../icons/rest.svg"
                    text: "Сделать выходным"
                    customColor: AppTheme.accentDanger
                    onClicked: {
                        backend.setDayType(root.targetDate, "weekend")
                        root.closeFromCell()
                    }
                }
                AppMenuRow {
                    visible: !root.menuIsHoliday
                    iconSource: "../icons/sparkle.svg"
                    text: "Сделать праздничным"
                    customColor: AppTheme.accentPurple
                    onClicked: {
                        backend.setDayType(root.targetDate, "holiday")
                        root.closeFromCell()
                    }
                }
            }
        }
    }

    // ============================================================
    // АНИМАЦИИ РАМКИ
    // ============================================================
    property real boxOpacity: 0
    Behavior on boxOpacity { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }

    ParallelAnimation {
        id: growAnim
        NumberAnimation { id: gx; target: box; property: "x";      duration: root.durGrow; easing.type: AppTheme.easeStandard }
        NumberAnimation { id: gy; target: box; property: "y";      duration: root.durGrow; easing.type: AppTheme.easeStandard }
        NumberAnimation { id: gw; target: box; property: "width";  duration: root.durGrow; easing.type: AppTheme.easeStandard }
        NumberAnimation { id: gh; target: box; property: "height"; duration: root.durGrow; easing.type: AppTheme.easeStandard }
        onStopped: {
            var cb = root._growCb
            root._growCb = null
            if (cb) cb()
        }
    }
    property var _growCb: null

    Timer { id: revealTimer; interval: 140; repeat: false; onTriggered: { menuContent.opacity = 1.0; menuContent.scale = 1.0 } }

    // ============================================================
    // ФОРМАТ ДАТЫ
    // ============================================================
    function fmtDate(dateString) {
        if (!dateString) return ""
        let parts = dateString.split("-")
        if (parts.length !== 3) return dateString
        let months = ["января", "февраля", "марта", "апреля", "мая", "июня",
                      "июля", "августа", "сентября", "октября", "ноября", "декабря"]
        return parseInt(parts[2], 10) + " " + months[parseInt(parts[1], 10) - 1] + " " + parts[0] + " г."
    }

    // ============================================================
    // ОТКРЫТИЕ ИЗ ЯЧЕЙКИ
    // ============================================================
    function openAtCell(cellItem, dateStr, isWeekend, isHoliday, hasDuties, hasComps) {
        if (dayDutyDialog.opened) dayDutyDialog.close()
        if (dayCompDialog.opened) dayCompDialog.close()
        if (dayInspector.opened) dayInspector.close()

        root.targetDate = dateStr
        root.menuIsWeekend = isWeekend
        root.menuIsHoliday = isHoliday
        root.menuHasDuties = hasDuties
        root.menuHasComps = hasComps

        let pt = cellItem.mapToItem(root, 0, 0)
        let margin = AppTheme.spaceL
        root.cellRect = Qt.rect(pt.x, pt.y, cellItem.width, cellItem.height)

        let menuW = Math.min(root.menuW, mainWindow.width - margin)
        let menuH = Math.min(menuContent.implicitHeight + 2 * AppTheme.spaceM, mainWindow.height - margin)

        // Появляемся у ячейки, но не вылезаем за окно
        let tX = Math.min(Math.max(pt.x, margin), mainWindow.width - menuW - margin)
        let tY = Math.min(Math.max(pt.y, margin), mainWindow.height - menuH - margin)

        // Сброс и подготовка
        root._growCb = null
        menuContent.opacity = 0
        menuContent.scale = 0.97

        // Сразу ставим финальный размер и включаем плавное появление
        // как у контекстных меню (прозрачность + масштаб).
        box.x = tX; box.y = tY
        box.width = menuW; box.height = menuH
        box.scale = 0.95
        root.mode = 1
        root.boxOpacity = 1
        box.scale = 1.0

        revealTimer.restart()
    }

    // ============================================================
    // ДОРАСТАНИЕ ДО СЛЕДУЮЩЕГО ОКНА
    // ============================================================
    function morphToDialog(dialogW, dialogH, done) {
        // Прячем содержимое меню — останется чистая рамка, которая растёт в диалог
        menuContent.opacity = 0
        revealTimer.stop()

        let margin = AppTheme.spaceL
        let w = Math.min(dialogW, mainWindow.width - 2 * margin)
        let h = Math.min(dialogH, mainWindow.height - 2 * margin)
        let x = margin + (mainWindow.width - 2 * margin - w) / 2
        let y = margin + (mainWindow.height - 2 * margin - h) / 2

        root.mode = 2

        // Сразу открываем целевое окно: оно проявляется во время роста
        // рамки, а не после — бесшовный переход.
        root._growCb = function() {
            root.boxOpacity = 0
            root.mode = 0
        }
        if (done) done(x, y, w, h)

        gx.to = x; gy.to = y; gw.to = w; gh.to = h
        growAnim.restart()
    }

    // ============================================================
    // ЗАКРЫТИЕ (действие выполнено / клик мимо / крестик)
    // ============================================================
    function closeFromCell() {
        root.boxOpacity = 0
        root.mode = 0
        revealTimer.stop()
        menuContent.opacity = 0
    }
}
