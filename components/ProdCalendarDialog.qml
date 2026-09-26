import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО «ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ» (2026 / 2027)
//
// Двенадцать месячных сеток с праздниками, переносами и
// предпраздничными днями, нормы рабочего времени под каждым
// месяцем и итог года. Данные — из официального производственного
// календаря (КонсультантПлюс), сверены с таблицей норм до часа.
// ============================================================
Popup {
    id: root

    property int calYear: 2026
    property var cal: ({})
    property var years: [2026, 2027]

    width: 900
    height: Math.min(
        (ApplicationWindow.window ? ApplicationWindow.window.height : 800) * 0.92,
        (ApplicationWindow.window ? ApplicationWindow.window.height : 800) - 60)
    z: AppTheme.zModal
    modal: true
    dim: true
    focus: true
    closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside

    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.96; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.96; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        color: AppTheme.bgModal
        radius: AppTheme.radiusModal
        border.color: AppTheme.borderDivider
        border.width: 1
        AppShadow { level: 4 }
    }

    function openYear(y) {
        calYear = y
        cal = backend.getProdCalendar(y)
        showCentered()
        grid.positionViewAtBeginning()
    }

    function showCentered() {
        if (ApplicationWindow.window) {
            root.x = (ApplicationWindow.window.width - root.width) / 2
            root.y = (ApplicationWindow.window.height - root.height) / 2
        }
        root.open()
    }

    contentItem: Item {
        anchors.fill: parent

        // ================= ШАПКА =================
        Item {
            id: headerBar
            anchors.top: parent.top; anchors.left: parent.left; anchors.right: parent.right
            height: 60

            Text {
                anchors.left: parent.left
                anchors.leftMargin: AppTheme.spaceL
                anchors.verticalCenter: parent.verticalCenter
                text: "Производственный календарь"
                color: AppTheme.textPrimary
                font.family: AppTheme.fontCondensed
                font.pixelSize: AppTheme.sizeH4
                font.weight: AppTheme.weightBold
            }

            // Переключатель годов
            Row {
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceM + 40
                anchors.verticalCenter: parent.verticalCenter
                spacing: AppTheme.spaceXS

                Repeater {
                    model: root.years
                    delegate: Rectangle {
                        width: 72; height: 34
                        radius: AppTheme.radiusMedium
                        color: modelData === root.calYear ? AppTheme.stateSelected
                               : (yHov.containsMouse ? AppTheme.stateHover : "transparent")
                        border.color: modelData === root.calYear ? "transparent" : AppTheme.borderDivider
                        border.width: 1
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                        Text {
                            anchors.centerIn: parent
                            text: modelData
                            color: modelData === root.calYear ? AppTheme.textOnSoft : AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                        }
                        MouseArea { id: yHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.openYear(modelData) }
                    }
                }
            }

            Rectangle {
                width: 32; height: 32; radius: AppTheme.radiusPill
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceM
                anchors.verticalCenter: parent.verticalCenter
                color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                IconImage { anchors.centerIn: parent; source: "../icons/close.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.textSecondary }
                MouseArea { id: closeHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.close() }
            }
        }

        Rectangle {
            anchors.left: parent.left; anchors.right: parent.right
            anchors.top: headerBar.bottom
            height: 1
            color: AppTheme.borderDivider
        }

        // ================= СЕТКА МЕСЯЦЕВ =================
        ScrollView {
            id: scrollArea
            anchors.top: headerBar.bottom
            anchors.topMargin: 1
            anchors.left: parent.left; anchors.right: parent.right
            anchors.bottom: footerBar.top
            clip: true
            ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

            GridView {
                id: grid
                objectName: "calGrid"
                anchors.fill: parent
                anchors.leftMargin: AppTheme.spaceL
                anchors.rightMargin: AppTheme.spaceL
                anchors.topMargin: AppTheme.spaceM
                anchors.bottomMargin: AppTheme.spaceM
                cellWidth: (width - AppTheme.spaceS * 2) / 3
                cellHeight: 290
                model: (root.cal && root.cal.months) ? root.cal.months : []
                boundsBehavior: Flickable.StopAtBounds

                delegate: Rectangle {
                    width: grid.cellWidth - AppTheme.spaceS
                    height: grid.cellHeight - AppTheme.spaceS
                    radius: AppTheme.radiusMedium
                    color: AppTheme.bgSurface
                    border.color: AppTheme.borderDivider
                    border.width: 1

                    Column {
                        anchors.fill: parent
                        anchors.margins: AppTheme.spaceS
                        spacing: AppTheme.spaceXS

                        // Название месяца + нормы
                        RowLayout {
                            width: parent.width
                            Text {
                                Layout.fillWidth: true
                                text: modelData.name
                                color: AppTheme.textPrimary
                                font.family: AppTheme.fontCondensed
                                font.pixelSize: AppTheme.sizeH5
                                font.weight: AppTheme.weightBold
                            }
                            Text {
                                text: modelData.workDays + " дн · " + modelData.h40 + " ч"
                                color: AppTheme.textTertiary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                font.weight: AppTheme.weightMedium
                            }
                        }

                        // Дни недели
                        Row {
                            width: parent.width
                            Repeater {
                                model: ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"]
                                Text {
                                    width: parent.width / 7
                                    horizontalAlignment: Text.AlignHCenter
                                    text: modelData
                                    color: AppTheme.textTertiary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: 10
                                    font.weight: AppTheme.weightBold
                                }
                            }
                        }

                        // Дни месяца
                        Grid {
                            width: parent.width
                            columns: 7
                            spacing: 2

                            Repeater {
                                // пустые клетки до первого дня + сами дни
                                model: {
                                    var cells = []
                                    for (var b = 0; b < modelData.firstWd; b++) cells.push(null)
                                    var list = modelData.cells || []
                                    for (var i = 0; i < list.length; i++) cells.push(list[i])
                                    return cells
                                }

                                delegate: Item {
                                    width: (grid.cellWidth - AppTheme.spaceS - AppTheme.spaceS * 2 - 12) / 7
                                    height: 24

                                    Rectangle {
                                        anchors.fill: parent
                                        radius: AppTheme.radiusSmall
                                        color: !modelData ? "transparent"
                                               : (modelData.hol ? AppTheme.bgDangerSoft
                                                  : (modelData.off ? AppTheme.bgPanel : AppTheme.bgCell))
                                        border.width: 1
                                        border.color: modelData && modelData.off && !modelData.hol ? AppTheme.borderDivider : "transparent"

                                        Text {
                                            anchors.centerIn: parent
                                            text: modelData ? modelData.d : ""
                                            color: !modelData ? "transparent"
                                                   : (modelData.hol ? AppTheme.accentDanger
                                                      : (modelData.off ? AppTheme.textSecondary : AppTheme.textPrimary))
                                            font.family: AppTheme.fontFamily
                                            font.pixelSize: AppTheme.sizeSmall
                                            font.weight: modelData && (modelData.hol || modelData.pre) ? AppTheme.weightBold : AppTheme.weightMedium
                                        }

                                        // Предпраздничный день: точка-звёздочка
                                        Rectangle {
                                            visible: modelData && modelData.pre && !modelData.off
                                            anchors.horizontalCenter: parent.horizontalCenter
                                            anchors.bottom: parent.bottom
                                            anchors.bottomMargin: 1
                                            width: 4; height: 4; radius: 2
                                            color: AppTheme.accentWarning
                                        }
                                    }
                                }
                            }
                        }

                        Item { width: 1; height: 1 }

                        // Нормы при 36- и 24-часовых неделях
                        Text {
                            width: parent.width
                            text: "36-час. неделя: " + modelData.h36 + " ч · 24-час.: " + modelData.h24 + " ч"
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: 10
                        }
                    }
                }
            }
        }

        // ================= ПОДВАЛ: ИТОГ + ЛЕГЕНДА =================
        Rectangle {
            id: footerBar
            anchors.left: parent.left; anchors.right: parent.right
            anchors.bottom: parent.bottom
            height: 74
            color: "transparent"

            Rectangle {
                anchors.top: parent.top
                anchors.left: parent.left; anchors.right: parent.right
                height: 1
                color: AppTheme.borderDivider
            }

            Column {
                anchors.fill: parent
                anchors.margins: AppTheme.spaceS
                spacing: AppTheme.spaceXXS

                Text {
                    width: parent.width
                    visible: Boolean(root.cal && root.cal.totals)
                    text: root.cal && root.cal.totals
                          ? root.calYear + " год: " + root.cal.totals.work + " рабочих дней · "
                            + root.cal.totals.h40 + " ч при 40-часовой неделе"
                          : ""
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                }
                Text {
                    width: parent.width
                    visible: Boolean(root.cal && root.cal.note)
                    text: (root.cal && root.cal.note ? root.cal.note + "  " : "")
                          + "Красным — нерабочие праздничные и перенесённые дни, серым — выходные, точкой отмечены предпраздничные (сокращённые на час)."
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    wrapMode: Text.WordWrap
                    elide: Text.ElideRight
                    maximumLineCount: 2
                }
            }
        }
    }
}
