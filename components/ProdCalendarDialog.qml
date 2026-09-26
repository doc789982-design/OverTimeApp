import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО «ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ» (2026 / 2027)
//
// Двенадцать месячных сеток. Красная звезда в углу ячейки —
// нерабочий ПРАЗДНИЧНЫЙ день по ст. 112 ТК РФ (работа в такой
// день оплачивается в двойном размере): ровно 14 дней в году,
// перенесённые выходные сюда не попадают. Серым — выходные дни,
// в том числе перенесённые (сб/вс и дни с переносов), оранжевой
// точкой — предпраздничные (сокращённые на час). Под каждым
// месяцем — полные нормы без сокращений. Внизу окна, всегда на
// виду: обозначения, справка о переносах и итог года. Прокрутка
// — плавная, с инерцией, как в браузере. Данные сверены с
// таблицей норм до часа.
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
            scroll.cancelWheelScroll()
            scroll.contentY = 0
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

            // Переключатель годов — настоящие кнопки:
            // выбранный год залит, второй — контурный
            Row {
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceM + 40
                anchors.verticalCenter: parent.verticalCenter
                spacing: AppTheme.spaceXS

                Repeater {
                    model: root.years
                    delegate: AppButton {
                        width: 88
                        text: modelData
                        variant: modelData === root.calYear ? "primary" : "secondary"
                        onClicked: root.openYear(modelData)
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
        // Плавная прокрутка с инерцией — как в браузере
        SmoothFlickable {
            id: scroll
            objectName: "calScroll"
            anchors.top: headerBar.bottom
            anchors.topMargin: 1
            anchors.left: parent.left; anchors.right: parent.right
            anchors.bottom: footerBar.top
            clip: true
            flickableDirection: Flickable.VerticalFlick
            boundsBehavior: Flickable.StopAtBounds
            contentWidth: width
            contentHeight: scrollColumn.height
            ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }

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
                                            height: 28
                                            radius: AppTheme.radiusSmall
                                            color: !modelData ? "transparent"
                                                   : (modelData.hol ? AppTheme.bgDangerSoft
                                                      : (modelData.off ? AppTheme.bgPanel : AppTheme.bgCell))

                                            // Праздничный день: красная звезда
                                            // в верхнем левом углу, чуть повёрнутая —
                                            // число остаётся свободным по центру
                                            Image {
                                                visible: modelData && modelData.hol
                                                anchors.top: parent.top
                                                anchors.left: parent.left
                                                anchors.topMargin: 1
                                                anchors.leftMargin: 2
                                                width: 12
                                                height: 12
                                                rotation: 18
                                                source: AppTheme.isDark
                                                        ? "../icons/sparkle_danger_dark.svg"
                                                        : "../icons/sparkle_danger_light.svg"
                                                fillMode: Image.PreserveAspectFit
                                                smooth: true
                                                mipmap: true
                                            }

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

                bottomPadding: AppTheme.spaceS
            }
        }

        // ============ ПОДВАЛ: ОБОЗНАЧЕНИЯ · СПРАВКА · ИТОГ ГОДА ============
        Rectangle {
            id: footerBar
            objectName: "calFooter"
            anchors.left: parent.left; anchors.right: parent.right
            anchors.bottom: parent.bottom
            height: footerCol.implicitHeight + AppTheme.spaceS * 2
            color: "transparent"

            Rectangle {
                anchors.top: parent.top
                anchors.left: parent.left; anchors.right: parent.right
                height: 1
                color: AppTheme.borderDivider
            }

            Column {
                id: footerCol
                anchors.left: parent.left; anchors.right: parent.right
                anchors.top: parent.top
                anchors.margins: AppTheme.spaceS
                spacing: AppTheme.spaceXS

                // Обозначения — всегда на виду, при любом положении прокрутки
                Flow {
                    width: parent.width
                    spacing: AppTheme.spaceL

                    Row {
                        spacing: AppTheme.spaceXS
                        Image {
                            anchors.verticalCenter: parent.verticalCenter
                            width: 15; height: 15
                            source: AppTheme.isDark
                                    ? "../icons/sparkle_danger_dark.svg"
                                    : "../icons/sparkle_danger_light.svg"
                            fillMode: Image.PreserveAspectFit
                            smooth: true
                            mipmap: true
                        }
                        Text {
                            anchors.verticalCenter: parent.verticalCenter
                            text: "нерабочий праздничный день — работа оплачивается в двойном размере"
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                    }

                    Row {
                        spacing: AppTheme.spaceXS
                        Rectangle {
                            anchors.verticalCenter: parent.verticalCenter
                            width: 16; height: 16
                            radius: AppTheme.radiusSmall
                            color: AppTheme.bgPanel
                            border.color: AppTheme.borderDivider
                            border.width: 1
                        }
                        Text {
                            anchors.verticalCenter: parent.verticalCenter
                            text: "выходные дни, в том числе перенесённые (субботы и воскресенья)"
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                    }

                    Row {
                        spacing: AppTheme.spaceXS
                        Item {
                            anchors.verticalCenter: parent.verticalCenter
                            width: 16; height: 16
                            Rectangle {
                                anchors.centerIn: parent
                                width: 16; height: 16
                                radius: AppTheme.radiusSmall
                                color: AppTheme.bgCell
                                border.color: AppTheme.borderDivider
                                border.width: 1
                            }
                            Rectangle {
                                anchors.horizontalCenter: parent.horizontalCenter
                                anchors.bottom: parent.bottom
                                anchors.bottomMargin: 2
                                width: 4; height: 4; radius: 2
                                color: AppTheme.accentWarning
                            }
                        }
                        Text {
                            anchors.verticalCenter: parent.verticalCenter
                            text: "предпраздничный — на час короче"
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                    }
                }

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

                // Справка о переносах — из производственного календаря
                Text {
                    width: parent.width
                    visible: Boolean(root.cal && root.cal.note && root.cal.note !== "")
                    text: (root.cal && root.cal.note) ? root.cal.note : ""
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
