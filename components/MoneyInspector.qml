import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects
import "."

// ============================================================
// ЖУРНАЛ ПРИКАЗОВ о денежной компенсации
//
// Приказ = один номер и дата, сколько бы сотрудников он ни
// охватывал: группе из 12 человек соответствует ОДНА строка
// с бейджем «×12». Клик раскрывает список получателей и суммы.
// Отсюда же — «Повторить приказ» (ведомость заполнится как в
// прошлый раз) и удаление приказа целиком.
// ============================================================
AppDialog {
    id: root
    width: 520
    heightFraction: 2 / 3
    title: "Приказы о денежной компенсации"
    acceptText: "Новый приказ"
    acceptVariant: "primary"
    rejectText: "Закрыть"

    signal requestNewOrder()

    onAccepted: root.requestNewOrder()

    // Совместимость со старым вызовом (из сводной панели)
    function show() { root.showCentered() }

    // Раскрытый приказ (order_no + order_date)
    property string expandedKey: ""

    function keyOf(order) { return order.order_no + "|" + order.order_date }

    Column {
        width: parent.width
        spacing: AppTheme.spaceS
        visible: backend.moneyOrders.length > 0

        Repeater {
            model: backend.moneyOrders

            Rectangle {
                width: parent.width
                height: {
                    let h = AppTheme.rowHeight + AppTheme.spaceS
                    if (root.keyOf(modelData) === root.expandedKey)
                        h += modelData.employees.length * 30 + AppTheme.spaceS
                    if (modelData.comment !== "") h += AppTheme.spaceL
                    return h
                }
                color: AppTheme.bgSurface
                border.color: root.keyOf(modelData) === root.expandedKey
                              ? AppTheme.borderFocus : AppTheme.borderDivider
                border.width: 1
                radius: AppTheme.radiusMedium
                Behavior on height { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }

                AppShadow { level: 1 }

                ColumnLayout {
                    anchors.fill: parent
                    anchors.margins: AppTheme.spaceM
                    anchors.rightMargin: AppTheme.cardActionReserve
                    spacing: AppTheme.spaceXXS

                    RowLayout {
                        Layout.fillWidth: true
                        spacing: AppTheme.spaceS

                        IconImage { source: "../icons/money.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.accentTeal }

                        Text {
                            text: modelData.order_no !== "" ? "Приказ № " + modelData.order_no : "Приказ"
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                        }

                        // Скольким сотрудникам — одним приказом
                        Rectangle {
                            visible: modelData.count > 1
                            width: countText.implicitWidth + 14
                            height: 20
                            radius: AppTheme.radiusPill
                            color: AppTheme.bgTealSoft
                            Text {
                                id: countText
                                anchors.centerIn: parent
                                text: "×" + modelData.count
                                color: AppTheme.accentTeal
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                font.weight: AppTheme.weightBold
                            }
                        }

                        Item { Layout.fillWidth: true }

                        Text {
                            text: {
                                let parts = []
                                if (modelData.hours > 0) parts.push(modelData.hours + " ч")
                                if (modelData.overtime > 0) parts.push(modelData.overtime + " ч сверх.")
                                if (modelData.days > 0) parts.push(modelData.days + " д")
                                return parts.length > 0 ? parts.join(" · ") : "—"
                            }
                            color: AppTheme.accentTeal
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                        }
                    }

                    RowLayout {
                        Layout.fillWidth: true
                        Text { text: modelData.date; color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall }
                        Text {
                            visible: modelData.count > 1
                            text: "получателей: " + modelData.count
                            color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                        }
                        Item { Layout.fillWidth: true }
                        Text {
                            visible: root.keyOf(modelData) === root.expandedKey
                            text: "свернуть ▲"
                            color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                        }
                        Text {
                            visible: root.keyOf(modelData) !== root.expandedKey && modelData.count > 1
                            text: "получатели ▼"
                            color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                        }
                    }

                    Text {
                        visible: modelData.comment !== ""
                        text: modelData.comment
                        color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                        elide: Text.ElideRight; Layout.fillWidth: true
                    }

                    // Список получателей (раскрыт)
                    Column {
                        visible: root.keyOf(modelData) === root.expandedKey
                        Layout.fillWidth: true
                        spacing: 2

                        Repeater {
                            model: modelData.employees

                            RowLayout {
                                width: parent.width
                                spacing: AppTheme.spaceS
                                Text {
                                    text: modelData.name
                                    color: AppTheme.textSecondary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeSmall
                                    elide: Text.ElideRight
                                    Layout.fillWidth: true
                                }
                                Text {
                                    text: {
                                        let parts = []
                                        if (modelData.hours > 0) parts.push(modelData.hours + " ч")
                                        if (modelData.overtime > 0) parts.push(modelData.overtime + " ч сверх.")
                                        if (modelData.days > 0) parts.push(modelData.days + " д")
                                        return parts.join(" · ")
                                    }
                                    color: AppTheme.textPrimary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeSmall
                                    font.weight: AppTheme.weightMedium
                                }
                            }
                        }
                    }
                }

                // Клик по карточке — раскрыть/свернуть получателей
                HoverHandler { cursorShape: Qt.PointingHandCursor }
                MouseArea {
                    anchors.fill: parent
                    hoverEnabled: true
                    cursorShape: Qt.PointingHandCursor
                    onClicked: {
                        root.expandedKey = root.keyOf(modelData) === root.expandedKey
                                         ? "" : root.keyOf(modelData)
                    }
                }

                // КНОПКИ СПРАВА
                Row {
                    anchors.right: parent.right
                    anchors.rightMargin: AppTheme.spaceS
                    anchors.verticalCenter: parent.verticalCenter
                    spacing: AppTheme.spaceXS

                    // Редактировать приказ
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: repHover.pressed ? AppTheme.statePress : (repHover.containsMouse ? AppTheme.stateHover : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/edit.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: repHover.containsMouse ? AppTheme.textPrimary : AppTheme.textTertiary }
                        HoverHandler { cursorShape: Qt.PointingHandCursor }
                        MouseArea {
                            id: repHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor
                            onClicked: { root.close(); mainWindow.editMoneyOrder(modelData) }
                        }
                        AppToolTip {
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.bottom: parent.top; anchors.bottomMargin: AppTheme.spaceXXS
                            text: "Редактировать приказ"; isVisible: repHover.containsMouse
                        }
                    }

                    // Удалить приказ
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: delMHover.pressed ? AppTheme.statePress : (delMHover.containsMouse ? AppTheme.bgDangerSoft : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/trash.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: delMHover.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary }
                        HoverHandler { cursorShape: Qt.PointingHandCursor }
                        MouseArea {
                            id: delMHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                mainWindow.askConfirm(
                                    "Удалить приказ?",
                                    "Будут удалены все записи приказа № " + modelData.order_no +
                                    " (" + modelData.count + " сотр.).\nЕсли передумаете — нажмите Ctrl+Z.",
                                    "Удалить",
                                    function() { backend.deleteMoneyOrder(modelData.order_no, modelData.order_date) }
                                )
                            }
                        }
                        AppToolTip {
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.bottom: parent.top; anchors.bottomMargin: AppTheme.spaceXXS
                            text: "Удалить приказ"; isVisible: delMHover.containsMouse
                        }
                    }
                }
            }
        }
    }

    Item {
        width: parent.width; implicitHeight: 200
        visible: backend.moneyOrders.length === 0
        Text { text: "В этом месяце приказов нет"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; anchors.centerIn: parent }
    }
}
