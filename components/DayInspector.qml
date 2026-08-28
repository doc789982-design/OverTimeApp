import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

// Окно дня («Открыть день») — в том же стиле, что и окна
// добавления дежурства / компенсации (AppDialog).
AppDialog {
    id: root
    width: 450

    property string targetDate: ""

    // Дата в заголовке, справа крестик закрытия — как у остальных окон.
    title: formatBeautifulDate(targetDate)
    acceptText: "Готово"
    rejectText: "Закрыть"
    showAccept: false
    showReject: true

    function formatBeautifulDate(dateString) {
        if (!dateString) return ""
        let parts = dateString.split("-")
        if (parts.length !== 3) return dateString
        let year = parts[0]; let month = parseInt(parts[1], 10); let day = parseInt(parts[2], 10)
        let months = ["января", "февраля", "марта", "апреля", "мая", "июня",
                      "июля", "августа", "сентября", "октября", "ноября", "декабря"]
        return day + " " + months[month - 1] + " " + year + " г."
    }

    Column {
        width: parent.width; spacing: AppTheme.spaceM
        visible: backend.dayDuties.length > 0

        Text {
            text: "ДЕЖУРСТВА"
            color: AppTheme.textTertiary; font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall; font.weight: AppTheme.weightBold; font.letterSpacing: 1
        }

        Repeater {
            model: backend.dayDuties
            // КАРТОЧКА ДЕЖУРСТВА
            Rectangle {
                width: parent.width; height: AppTheme.rowHeight + AppTheme.spaceS
                color: AppTheme.bgSurface
                radius: AppTheme.radiusMedium
                border.color: AppTheme.borderDivider; border.width: 1

                // Тень-картинка вместо вычисляемой (Level 1)
                AppShadow { level: 1 }

                RowLayout {
                    anchors.fill: parent; anchors.margins: AppTheme.spaceM; anchors.rightMargin: AppTheme.cardActionReserve
                    spacing: AppTheme.spaceM
                    IconImage { source: "../icons/clock.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.accentBrand }
                    Text {
                        text: modelData.start + " — " + modelData.end
                        color: AppTheme.textPrimary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; font.weight: AppTheme.weightBold
                        Layout.preferredWidth: 100
                    }
                    Text {
                        text: modelData.comment
                        color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody
                        elide: Text.ElideRight; Layout.fillWidth: true
                    }
                }

                // КНОПКИ СПРАВА
                Row {
                    anchors.right: parent.right; anchors.rightMargin: AppTheme.spaceS
                    anchors.verticalCenter: parent.verticalCenter; spacing: AppTheme.spaceXS

                    // Редактировать
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: editDutyHover.pressed ? AppTheme.statePress : (editDutyHover.containsMouse ? AppTheme.stateHover : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/edit.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: editDutyHover.containsMouse ? AppTheme.textPrimary : AppTheme.textTertiary }
                        MouseArea { id: editDutyHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: dayDutyDialog.openForDutyEdit(modelData, root.targetDate, parent, 0, 0) }
                    }

                    // Удалить
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: delDutyHover.pressed ? AppTheme.statePress : (delDutyHover.containsMouse ? AppTheme.bgDangerSoft : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/trash.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: delDutyHover.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary }
                        MouseArea { id: delDutyHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: { mainWindow.explodeAndDelete(root.targetDate, "duty", modelData.id, function() { backend.deleteDuty(modelData.id, root.targetDate) }) } }
                    }
                }
            }
        }
    }

    Item { width: parent.width; implicitHeight: AppTheme.spaceL; visible: backend.dayDuties.length > 0 && backend.dayComps.length > 0 }

    Column {
        width: parent.width; spacing: AppTheme.spaceM
        visible: backend.dayComps.length > 0

        Text {
            text: "КОМПЕНСАЦИИ"
            color: AppTheme.textTertiary; font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall; font.weight: AppTheme.weightBold; font.letterSpacing: 1
        }

        Repeater {
            model: backend.dayComps
            // КАРТОЧКА КОМПЕНСАЦИИ
            Rectangle {
                width: parent.width; height: AppTheme.rowHeight + AppTheme.spaceS
                color: AppTheme.bgSurface
                radius: AppTheme.radiusMedium
                border.color: AppTheme.borderDivider; border.width: 1

                // Тень-картинка вместо вычисляемой (Level 1)
                AppShadow { level: 1 }

                RowLayout {
                    anchors.fill: parent; anchors.margins: AppTheme.spaceM; anchors.rightMargin: AppTheme.cardActionReserve
                    spacing: AppTheme.spaceM

                    IconImage {
                        source: "../icons/rest.svg"
                        width: AppTheme.iconMedium
                        height: AppTheme.iconMedium
                        color: modelData.unit === "overtime"
                               ? AppTheme.accentWarning
                               : AppTheme.accentTeal
                    }

                    Column {
                        spacing: 2
                        Layout.preferredWidth: 120

                        Text {
                            text: modelData.type
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                        }

                        Text {
                            text: modelData.amount
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                        }
                    }

                    Text {
                        text: modelData.comment
                        color: AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        elide: Text.ElideRight
                        Layout.fillWidth: true
                    }
                }

                // КНОПКИ СПРАВА
                Row {
                    anchors.right: parent.right; anchors.rightMargin: AppTheme.spaceS
                    anchors.verticalCenter: parent.verticalCenter; spacing: AppTheme.spaceXS

                    // Редактировать
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: editCompHover.pressed ? AppTheme.statePress : (editCompHover.containsMouse ? AppTheme.stateHover : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/edit.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: editCompHover.containsMouse ? AppTheme.textPrimary : AppTheme.textTertiary }
                        MouseArea { id: editCompHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: dayCompDialog.openForCompEdit(modelData, root.targetDate, parent, 0, 0) }
                    }

                    // Удалить
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: delCompHover.pressed ? AppTheme.statePress : (delCompHover.containsMouse ? AppTheme.bgDangerSoft : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/trash.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: delCompHover.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary }
                        MouseArea { id: delCompHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: { mainWindow.explodeAndDelete(root.targetDate, "comp", modelData.id, function() { backend.deleteCompensation(modelData.id, root.targetDate) }) } }
                    }
                }
            }
        }
    }

    Item {
        width: parent.width; implicitHeight: 200
        visible: backend.dayDuties.length === 0 && backend.dayComps.length === 0
        Text { text: "В этот день ничего не назначено."; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; anchors.centerIn: parent }
    }
}
