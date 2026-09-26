import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

// ============================================================
// МЕНЮ ДНЯ (открывается ЛКМ и ПКМ по ячейке календаря)
//
// Настоящее контекстное меню на той же базе, что и меню по ПКМ
// (Qt Quick Menu): та же рамка, тень, плавное появление
// (прозрачность + масштаб), закрытие по клику вне меню.
// Внутри — шапка с датой, круглые статусы К/Б/О и пункты действий.
// ============================================================
Menu {
    id: root

    z: AppTheme.zDropdown
    transformOrigin: Item.TopLeft
    padding: 0
    topPadding: AppTheme.spaceXS
    bottomPadding: AppTheme.spaceXS
    leftPadding: AppTheme.spaceXXS
    rightPadding: AppTheme.spaceXXS

    // Фиксированная ширина меню (как у старых меню дня): пункты и шапка
    // выравниваются по ней, а не по авто-размеру содержимого.
    width: 240

    // Параметры для открывающего дня
    property string targetDate: ""
    property bool menuIsWeekend: false
    property bool menuIsHoliday: false
    property bool menuHasDuties: false
    property bool menuHasComps: false

    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.96; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.96; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        implicitWidth: 240
        color: AppTheme.bgElevated
        radius: AppTheme.radiusMedium
        border.color: AppTheme.borderDivider
        border.width: 1
        AppShadow { level: 3 }
    }

    // ============================================================
    // ШАПКА: дата + закрыть
    // ============================================================
    MenuItem {
        id: headerItem
        implicitWidth: 232
        implicitHeight: 44
        background: Rectangle { color: "transparent" }

        contentItem: Item {
            implicitWidth: 232
            implicitHeight: 44

            Text {
                id: dateLabel
                anchors.left: parent.left
                anchors.leftMargin: AppTheme.spaceS
                anchors.verticalCenter: parent.verticalCenter
                anchors.right: parent.right
                anchors.rightMargin: 44
                elide: Text.ElideRight
                text: root.fmtDate(root.targetDate)
                color: AppTheme.textPrimary
                font.family: AppTheme.fontCondensed
                font.pixelSize: AppTheme.sizeH4
                font.weight: AppTheme.weightBold
                verticalAlignment: Text.AlignVCenter
            }

            Rectangle {
                width: 28; height: 28; radius: AppTheme.radiusPill
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceXS
                anchors.verticalCenter: parent.verticalCenter
                color: closeHov.pressed ? AppTheme.statePress
                      : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                IconImage {
                    anchors.centerIn: parent
                    source: "../icons/close.svg"
                    width: AppTheme.iconMedium
                    height: AppTheme.iconMedium
                    color: AppTheme.textSecondary
                }
                MouseArea {
                    id: closeHov
                    anchors.fill: parent
                    hoverEnabled: true
                    cursorShape: Qt.PointingHandCursor
                    onClicked: root.close()
                }
            }
        }
    }

    // ============================================================
    // БЫСТРЫЕ СТАТУСЫ: К / Б / О / сброс (круглые)
    // ============================================================
    MenuItem {
        id: statusItem
        implicitWidth: 232
        implicitHeight: 46
        background: Rectangle { color: "transparent" }

        contentItem: Item {
            implicitWidth: 232
            implicitHeight: 46

            Row {
                anchors.left: parent.left
                anchors.leftMargin: AppTheme.spaceS
                anchors.verticalCenter: parent.verticalCenter
                spacing: AppTheme.spaceS

                Repeater {
                    model: [
                        { "t": "К", "tool": "Командировка", "c": AppTheme.accentPurple, "st": "К" },
                        { "t": "Б", "tool": "Больничный",     "c": AppTheme.accentDanger, "st": "Б" },
                        { "t": "О", "tool": "Отпуск",         "c": AppTheme.accentWarning, "st": "О" }
                    ]
                    Rectangle {
                        width: 34; height: 34; radius: AppTheme.radiusPill
                        color: st.pressed ? AppTheme.statePress
                              : (st.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
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
                            id: st
                            anchors.fill: parent
                            hoverEnabled: true
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                backend.setDayStatus(root.targetDate, modelData.st)
                                root.close()
                            }
                        }
                        AppToolTip {
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.bottom: parent.top
                            anchors.bottomMargin: AppTheme.spaceXS
                            text: modelData.tool
                            isVisible: st.containsMouse
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
                            root.close()
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
        }
    }

    AppMenuSeparator { }

    // ============================================================
    // ПУНКТЫ ДЕЙСТВИЙ
    // ============================================================
    AppMenuItem {
        iconSource: "../icons/edit.svg"
        text: "Открыть день"
        onClicked: {
            backend.loadDayDetails(root.targetDate)
            dayInspector.targetDate = root.targetDate
            root.close()
            dayInspector.showCentered()
        }
    }

    AppMenuItem {
        iconSource: "../icons/clock.svg"
        text: "Добавить дежурство"
        showDelete: root.menuHasDuties
        onClicked: {
            root.close()
            dayDutyDialog.prepareForDuty(root.targetDate)
            dayDutyDialog.showCentered()
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
            root.close()
        }
    }

    AppMenuItem {
        iconSource: "../icons/rest.svg"
        text: "Добавить компенсацию"
        showDelete: root.menuHasComps
        onClicked: {
            root.close()
            dayCompDialog.prepareForComp(root.targetDate)
            dayCompDialog.showCentered()
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
            root.close()
        }
    }

    AppMenuSeparator { }

    AppMenuItem {
        visible: root.menuIsWeekend || root.menuIsHoliday
        iconSource: "../icons/calendar.svg"
        text: "Сделать рабочим"
        customColor: AppTheme.accentTeal
        onClicked: {
            backend.setDayType(root.targetDate, "work")
            root.close()
        }
    }
    AppMenuItem {
        visible: !root.menuIsWeekend
        iconSource: "../icons/rest.svg"
        text: "Сделать выходным"
        customColor: AppTheme.accentDanger
        onClicked: {
            backend.setDayType(root.targetDate, "weekend")
            root.close()
        }
    }
    AppMenuItem {
        visible: !root.menuIsHoliday
        iconSource: "../icons/sparkle.svg"
        text: "Сделать праздничным"
        customColor: AppTheme.accentPurple
        onClicked: {
            backend.setDayType(root.targetDate, "holiday")
            root.close()
        }
    }

    // ============================================================
    // ОТКРЫТИЕ ИЗ ЯЧЕЙКИ ДНЯ
    // ============================================================
    function openFromCell(cellItem, dateStr, isWeekend, isHoliday, hasDuties, hasComps) {
        if (dayDutyDialog.opened) dayDutyDialog.close()
        if (dayCompDialog.opened) dayCompDialog.close()
        if (dayInspector.opened) dayInspector.close()

        root.targetDate = dateStr
        root.menuIsWeekend = isWeekend
        root.menuIsHoliday = isHoliday
        root.menuHasDuties = hasDuties
        root.menuHasComps = hasComps

        let pt = cellItem.mapToItem(null, 0, 0)
        root.popup(Math.round(pt.x), Math.round(pt.y))
    }

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
}
