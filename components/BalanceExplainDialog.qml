import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО «ПОЧЕМУ ТАКАЯ ЦИФРА?»
//
// Открывается нажатием на число в панели «Итоги» и показывает
// расчёт человеческим списком: остаток на начало, дежурства,
// давшие начисления, компенсации и остаток на конец периода.
// Прошлый год не прячется в скобку, а показан отдельной строкой.
// ============================================================
AppModalShell {
    id: root
    closeOnOutside: true

    property var info: ({})
    property string bucket: "hours"
    property string scope: "month"

    width: Math.min(880, ApplicationWindow.window ? ApplicationWindow.window.width - 48 : 880)
    height: Math.min(760, ApplicationWindow.window ? ApplicationWindow.window.height - 32 : 760)

    function openExplain(newBucket, newScope) {
        root.bucket = newBucket
        root.scope = newScope
        root.info = backend.getBalanceExplanation(newBucket, newScope)
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

    function rowsCount(rows) {
        return rows && rows.length !== undefined ? rows.length : 0
    }

    component LedgerRow: Rectangle {
        id: rowRoot
        property var rowData: ({})
        width: parent ? parent.width : 0
        height: Math.max(52, rowLayout.implicitHeight + AppTheme.spaceM)
        radius: AppTheme.radiusMedium
        color: AppTheme.bgSurface
        border.color: AppTheme.borderDivider
        border.width: 1

        RowLayout {
            id: rowLayout
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.verticalCenter: parent.verticalCenter
            anchors.leftMargin: AppTheme.spaceM
            anchors.rightMargin: AppTheme.spaceM
            spacing: AppTheme.spaceM

            Text {
                Layout.preferredWidth: 92
                Layout.alignment: Qt.AlignTop
                visible: (rowRoot.rowData.date_text || "") !== ""
                text: rowRoot.rowData.date_text || ""
                color: AppTheme.textTertiary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightMedium
                wrapMode: Text.WordWrap
            }

            ColumnLayout {
                Layout.fillWidth: true
                Layout.alignment: Qt.AlignVCenter
                spacing: 2

                Text {
                    Layout.fillWidth: true
                    text: rowRoot.rowData.title || ""
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                    wrapMode: Text.WordWrap
                }
                Text {
                    Layout.fillWidth: true
                    visible: (rowRoot.rowData.detail || "") !== ""
                    text: rowRoot.rowData.detail || ""
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    wrapMode: Text.WordWrap
                }
            }

            Text {
                Layout.preferredWidth: 118
                Layout.alignment: Qt.AlignVCenter
                horizontalAlignment: Text.AlignRight
                text: rowRoot.rowData.amount_text || ""
                color: rowRoot.rowData.kind === "adjustment"
                       ? AppTheme.textSecondary
                       : ((rowRoot.rowData.amount_value || 0) < 0
                          ? AppTheme.textDanger
                          : ((rowRoot.rowData.amount_value || 0) > 0
                             ? AppTheme.accentSuccess
                             : AppTheme.textPrimary))
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                font.features: {"tnum": 1}
                font.weight: AppTheme.weightBold
                wrapMode: Text.WordWrap
            }
        }
    }

    contentItem: Item {
        anchors.fill: parent

        AppDialogHeader {
            id: headerBar
            anchors.top: parent.top
            anchors.left: parent.left
            anchors.right: parent.right
            title: "Почему такая цифра?"
            onCloseRequested: root.close()
        }
        Rectangle {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: headerBar.bottom
            height: 1
            color: AppTheme.borderDivider
        }

        SmoothFlickable {
            id: scroll
            anchors.top: headerBar.bottom
            anchors.topMargin: 1
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.bottom: footerBar.top
            clip: true
            flickableDirection: Flickable.VerticalFlick
            boundsBehavior: Flickable.StopAtBounds
            contentWidth: width
            contentHeight: bodyCol.height + AppTheme.spaceL
            ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }

            Column {
                id: bodyCol
                width: scroll.width - AppTheme.spaceL * 2
                x: AppTheme.spaceL
                y: AppTheme.spaceM
                spacing: AppTheme.spaceM

                Column {
                    width: parent.width
                    spacing: AppTheme.spaceXXS

                    Text {
                        width: parent.width
                        text: root.info.title || ""
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontCondensed
                        font.pixelSize: AppTheme.sizeH3
                        font.weight: AppTheme.weightBold
                        wrapMode: Text.WordWrap
                    }
                    Text {
                        width: parent.width
                        text: root.info.subtitle || ""
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        wrapMode: Text.WordWrap
                    }
                }

                Row {
                    width: parent.width
                    spacing: AppTheme.spaceXS

                    Repeater {
                        model: [
                            { "key": "hours", "label": "Ночные" },
                            { "key": "days", "label": "Дни" },
                            { "key": "overtime", "label": "Сверхурочные" },
                            { "key": "total", "label": "Всего" }
                        ]
                        delegate: AppButton {
                            width: (bodyCol.width - AppTheme.spaceXS * 3) / 4
                            height: 36
                            text: modelData.label
                            variant: root.bucket === modelData.key ? "primary" : "secondary"
                            onClicked: root.openExplain(modelData.key, root.scope)
                        }
                    }
                }

                Rectangle {
                    width: parent.width
                    height: formulaCol.implicitHeight + AppTheme.spaceM * 2
                    radius: AppTheme.radiusMedium
                    color: AppTheme.bgSurface
                    border.color: AppTheme.borderDivider
                    border.width: 1
                    AppShadow { level: 1 }

                    Column {
                        id: formulaCol
                        anchors.left: parent.left
                        anchors.right: parent.right
                        anchors.top: parent.top
                        anchors.margins: AppTheme.spaceM
                        spacing: AppTheme.spaceXS

                        RowLayout {
                            width: parent.width
                            spacing: AppTheme.spaceM
                            Text {
                                Layout.fillWidth: true
                                text: "Остаток на конец периода"
                                color: AppTheme.textTertiary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                font.weight: AppTheme.weightBold
                            }
                            Text {
                                text: root.info.result_text || "—"
                                color: AppTheme.textPrimary
                                font.family: AppTheme.fontCondensed
                                font.pixelSize: AppTheme.sizeH4
                                font.features: {"tnum": 1}
                                font.weight: AppTheme.weightBold
                            }
                        }
                        Text {
                            width: parent.width
                            visible: (root.info.current_formula || "") !== ""
                            text: root.info.current_formula || ""
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            wrapMode: Text.WordWrap
                        }
                        Text {
                            width: parent.width
                            visible: (root.info.prev_formula || "") !== ""
                            text: root.info.prev_formula || ""
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            wrapMode: Text.WordWrap
                        }
                    }
                }

                // ── Всего: формула перевода часов и дней в дни ──
                Column {
                    width: parent.width
                    visible: root.rowsCount(root.info.total_rows) > 0
                    spacing: AppTheme.spaceS

                    Text {
                        width: parent.width
                        text: "Как посчитано"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBodyLarge
                        font.weight: AppTheme.weightBold
                    }
                    Repeater {
                        model: root.info.total_rows || []
                        delegate: LedgerRow { rowData: modelData }
                    }
                }

                // ── Начисления по дежурствам ──
                Column {
                    width: parent.width
                    visible: root.bucket !== "total"
                    spacing: AppTheme.spaceS

                    Text {
                        width: parent.width
                        text: "Начисления"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBodyLarge
                        font.weight: AppTheme.weightBold
                    }
                    Text {
                        width: parent.width
                        visible: root.rowsCount(root.info.acc_rows) === 0
                        text: "Начислений за этот период не было."
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        wrapMode: Text.WordWrap
                    }
                    Repeater {
                        model: root.info.acc_rows || []
                        delegate: LedgerRow { rowData: modelData }
                    }
                }

                // ── Компенсации и пересчёты между копилками ──
                Column {
                    width: parent.width
                    visible: root.bucket !== "total"
                    spacing: AppTheme.spaceS

                    Text {
                        width: parent.width
                        text: "Компенсации"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBodyLarge
                        font.weight: AppTheme.weightBold
                    }
                    Text {
                        width: parent.width
                        visible: root.rowsCount(root.info.comp_rows) === 0
                        text: "Компенсаций за этот период не было."
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        wrapMode: Text.WordWrap
                    }
                    Repeater {
                        model: root.info.comp_rows || []
                        delegate: LedgerRow { rowData: modelData }
                    }
                }

                Text {
                    width: parent.width
                    visible: (root.info.note || "") !== ""
                    text: root.info.note || ""
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    wrapMode: Text.WordWrap
                }
            }
        }

        Item {
            id: footerBar
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.bottom: parent.bottom
            height: 66

            Rectangle {
                anchors.left: parent.left
                anchors.right: parent.right
                anchors.top: parent.top
                height: 1
                color: AppTheme.borderDivider
            }
            AppButton {
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceL
                anchors.verticalCenter: parent.verticalCenter
                width: 128
                height: 38
                text: "Понятно"
                variant: "primary"
                onClicked: root.close()
            }
        }
    }
}
