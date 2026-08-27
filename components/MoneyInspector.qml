import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

AppSidePanel {
    id: root
    width: 450 
    title: "Приказы и выплаты"
    
    signal requestAddMoneyDialog()

    AppButton {
        text: "+ Добавить выплату"
        width: parent.width
        variant: "success" 
        onClicked: root.requestAddMoneyDialog()
    }

    Item { width: parent.width; implicitHeight: AppTheme.spaceM; visible: backend.moneyComps.length > 0 }

    Column {
        width: parent.width
        spacing: AppTheme.spaceM
        visible: backend.moneyComps.length > 0

        Repeater {
            model: backend.moneyComps

            Rectangle {
                width: parent.width
                height: AppTheme.rowHeight + AppTheme.spaceL + (modelData.comment !== "" ? AppTheme.spaceL : 0) 
                color: AppTheme.bgSurface 
                border.color: AppTheme.borderDivider 
                border.width: 1                  
                radius: AppTheme.radiusMedium
                
                // Тень-картинка вместо вычисляемой (Level 1)
                AppShadow { level: 1 }
                
                ColumnLayout {
                    anchors.fill: parent; anchors.margins: AppTheme.spaceM; anchors.rightMargin: AppTheme.cardActionReserve; spacing: AppTheme.spaceXXS
                    
                    RowLayout {
                        Layout.fillWidth: true; spacing: AppTheme.spaceS
                        IconImage { source: "../icons/money.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.accentTeal }
                        Text { text: modelData.type; color: AppTheme.textPrimary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; font.weight: AppTheme.weightBold; Layout.fillWidth: true }
                        Text { text: modelData.amount; color: AppTheme.accentTeal; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; font.weight: AppTheme.weightBold }
                    }
                    
                    RowLayout {
                        Layout.fillWidth: true
                        Text { text: modelData.date; color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall }
                        Item { Layout.fillWidth: true } 
                        Text { text: "№ " + modelData.order_no; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall }
                    }

                    Text { 
                        visible: modelData.comment !== ""
                        text: modelData.comment
                        color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                        elide: Text.ElideRight; Layout.fillWidth: true; Layout.topMargin: AppTheme.spaceXS
                    }
                }
                
                // КНОПКИ СПРАВА
                Row {
                    anchors.right: parent.right; anchors.rightMargin: AppTheme.spaceS
                    anchors.verticalCenter: parent.verticalCenter; spacing: AppTheme.spaceXS
                    
                    // Редактировать
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: editMHover.pressed ? AppTheme.statePress : (editMHover.containsMouse ? AppTheme.stateHover : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/edit.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: editMHover.containsMouse ? AppTheme.textPrimary : AppTheme.textTertiary }
                        MouseArea { id: editMHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: moneyDialog.openEdit(modelData, parent, 0, 0) }
                    }
                    
                    // Удалить
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: delMHover.pressed ? AppTheme.statePress : (delMHover.containsMouse ? AppTheme.bgDangerSoft : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/trash.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: delMHover.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary }
                        MouseArea { id: delMHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: backend.deleteMoneyComp(modelData.id) }
                    }
                }
            }
        }
    }
    
    Item {
        width: parent.width; implicitHeight: 200 
        visible: backend.moneyComps.length === 0
        Text { text: "Нет выплат в этом месяце"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; anchors.centerIn: parent }
    }
}