import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО «ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ» (2026 / 2027)
//
// Двенадцать месячных сеток: красным — праздничные и перенесённые
// дни, серым — выходные, точкой отмечены предпраздничные
// (сокращённые на час). Под каждым месяцем — полные нормы без
// сокращений. Пояснения и переносы — внизу, прокручиваются вместе
// с календарём. Данные сверены с таблицей норм до часа.
// ============================================================
Popup {
    id: root

    property int calYear: 2026
    property var cal: ({})
    property var years: [2026, 2027]

    width: 940
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
        Qt.callLater(function() {
            if (scroll.flickableItem) scroll.flickableItem.contentY = 0
        })
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
                        width: 74; height: 34
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

        // ================= ПРОКРУЧИВАЕМЫЙ КОНТЕНТ =================
        ScrollView {
            id: scroll
            objectName: "calScroll"
            anchors.top: headerBar.bottom
            anchors.topMargin: 1
            anchors.left: parent.left; anchors.right: parent.right
            anchors.bottom: footerBar.top
            clip: true
            ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

            Column {
                id: scrollColumn
                width: scroll.width - AppTheme.spaceL * 2
                x: AppTheme.spaceL
                spacing: AppTheme.spaceM
                topPadding: AppTheme.spaceM

                // ── СЕТКА МЕСЯЦЕВ ──
                Flow {
                    id: monthsFlow
                    width: parent.width
                    spacing: AppTheme.spaceS

                    Repeater {
                        model: (root.cal && root.cal.months) ? root.cal.months : []

                        delegate: Rectangle {
                            id: monthCard
                            objectName: "monthCard"
                            width: (monthsFlow.width - monthsFlow.spacing * 2) / 3
                            height: 306
                            radius: AppTheme.radiusMedium
                            color: AppTheme.bgSurface
                            border.color: AppTheme.borderDivider
                            border.width: 1

                            readonly property real cellW: (width - AppTheme.spaceS * 2 - 2 * 6) / 7

                            Column {
                                anchors.fill: parent
                                anchors.margins: AppTheme.spaceS
                                spacing: 5

                                // Название месяца + рабочие дни
                                RowLayout {
                                    width: parent.width
                                    Text {
                                        Layout.fillWidth: true
                                        text: modelData.name
                                        color: AppTheme.textPrimary
                                        font.family: AppTheme.fontCondensed
                                        font.pixelSize: AppTheme.sizeH5
                                        font.weight: AppTheme.weightBold
                                        elide: Text.ElideRight
                                    }
                                    Text {
                                        text: modelData.workDays + " рабочих дней"
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
                                            width: monthCard.cellW
                                            horizontalAlignment: Text.AlignHCenter
                                            text: modelData
                                            color: index > 4 ? AppTheme.accentDanger : AppTheme.textTertiary
                                            opacity: index > 4 ? 0.55 : 1
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
                                        model: {
                                            var cells = []
                                            for (var b = 0; b < modelData.firstWd; b++) cells.push(null)
                                            var list = modelData.cells || []
                                            for (var i = 0; i < list.length; i++) cells.push(list[i])
                                            return cells
                                        }

                                        delegate: Rectangle {
                                            width: monthCard.cellW
                                            height: 25
                                            radius: AppTheme.radiusSmall
                                            color: !modelData ? "transparent"
                                                   : (modelData.hol ? AppTheme.bgDangerSoft
                                                      : (modelData.off ? AppTheme.bgPanel : AppTheme.bgCell))

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

                                Item { width: 1; height: 1 }

                                // Нормы месяца — полными словами
                                Text {
                                    width: parent.width
                                    text: "Часы за месяц: " + modelData.h40 + " при 40-часовой неделе, "
                                          + modelData.h36 + " при 36-часовой, "
                                          + modelData.h24 + " при 24-часовой."
                                    color: AppTheme.textTertiary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeSmall
                                    wrapMode: Text.WordWrap
                                    lineHeight: 1.25
                                }
                            }
                        }
                    }
                }

                // ── ПОЯСНЕНИЯ (прокручиваются вместе с календарём) ──
                Rectangle {
                    width: parent.width
                    height: legendCol.implicitHeight + AppTheme.spaceM * 2
                    radius: AppTheme.radiusMedium
                    color: AppTheme.bgSurface
                    border.color: AppTheme.borderDivider
                    border.width: 1

                    Column {
                        id: legendCol
                        anchors.left: parent.left
                        anchors.right: parent.right
                        anchors.margins: AppTheme.spaceM
                        anchors.verticalCenter: parent.verticalCenter
                        spacing: AppTheme.spaceXS

                        Text {
                            width: parent.width
                            text: "ОБОЗНАЧЕНИЯ"
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.letterSpacing: 1
                            font.weight: AppTheme.weightBold
                        }

                        Row {
                            spacing: AppTheme.spaceXS
                            Rectangle { width: 18; height: 18; radius: AppTheme.radiusSmall; color: AppTheme.bgDangerSoft; border.width: 1; border.color: AppTheme.accentDanger; anchors.top: parent.top }
                            Text {
                                width: legendCol.width - 18 - AppTheme.spaceXS
                                text: "красным — нерабочие праздничные дни и выходные, перенесённые на другие дни;"
                                color: AppTheme.textSecondary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                wrapMode: Text.WordWrap
                            }
                        }
                        Row {
                            spacing: AppTheme.spaceXS
                            Rectangle { width: 18; height: 18; radius: AppTheme.radiusSmall; color: AppTheme.bgPanel; border.width: 1; border.color: AppTheme.borderDivider; anchors.top: parent.top }
                            Text {
                                width: legendCol.width - 18 - AppTheme.spaceXS
                                text: "серым — обычные выходные дни (субботы и воскресенья);"
                                color: AppTheme.textSecondary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                wrapMode: Text.WordWrap
                            }
                        }
                        Row {
                            spacing: AppTheme.spaceXS
                            Item {
                                width: 18; height: 18
                                Rectangle { anchors.centerIn: parent; width: 18; height: 18; radius: AppTheme.radiusSmall; color: AppTheme.bgCell }
                                Rectangle { anchors.horizontalCenter: parent.horizontalCenter; anchors.bottom: parent.bottom; anchors.bottomMargin: 2; width: 4; height: 4; radius: 2; color: AppTheme.accentWarning }
                            }
                            Text {
                                width: legendCol.width - 18 - AppTheme.spaceXS
                                text: "оранжевая точка — предпраздничный день, служба сокращена на один час."
                                color: AppTheme.textSecondary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                wrapMode: Text.WordWrap
                            }
                        }
                    }
                }

                // ── ПЕРЕНОСЫ ВЫХОДНЫХ ──
                Rectangle {
                    width: parent.width
                    height: noteText.implicitHeight + AppTheme.spaceM * 2
                    radius: AppTheme.radiusMedium
                    color: AppTheme.bgSurface
                    border.color: AppTheme.borderDivider
                    border.width: 1

                    Column {
                        anchors.left: parent.left
                        anchors.right: parent.right
                        anchors.margins: AppTheme.spaceM
                        anchors.verticalCenter: parent.verticalCenter
                        spacing: AppTheme.spaceXXS

                        Text {
                            width: parent.width
                            text: "ПЕРЕНОСЫ ВЫХОДНЫХ ДНЕЙ"
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.letterSpacing: 1
                            font.weight: AppTheme.weightBold
                        }
                        Text {
                            id: noteText
                            width: parent.width
                            text: (root.cal && root.cal.note) ? root.cal.note : ""
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            wrapMode: Text.WordWrap
                            lineHeight: 1.3
                        }
                    }
                }

                bottomPadding: AppTheme.spaceS
            }
        }

        // ================= ПОДВАЛ: ИТОГ ГОДА =================
        Rectangle {
            id: footerBar
            anchors.left: parent.left; anchors.right: parent.right
            anchors.bottom: parent.bottom
            height: 66
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
                          ? root.calYear + " год: " + root.cal.totals.work + " рабочих дней и "
                            + root.cal.totals.h40 + " часов при 40-часовой рабочей неделе"
                          : ""
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                    elide: Text.ElideRight
                }
                Text {
                    width: parent.width
                    visible: Boolean(root.cal && root.cal.totals && root.cal.totals.h36)
                    text: root.cal && root.cal.totals && root.cal.totals.h36
                          ? "При 36-часовой неделе — " + root.cal.totals.h36 + " часов, при 24-часовой — "
                            + root.cal.totals.h24 + " часа."
                          : ""
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    elide: Text.ElideRight
                }
            }
        }
    }
}
