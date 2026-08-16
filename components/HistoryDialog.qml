import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

AppSidePanel {
    id: root
    width: 450

    property int targetEmpId: 0
    property string targetEmpName: ""

    title: "История переводов"

    Text {
        text: "Сотрудник: " + root.targetEmpName
        color: AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        width: parent.width
        wrapMode: Text.WordWrap
    }

    AppButton {
        text: "+ Добавить запись"
        width: parent.width
        variant: "secondary" 
        onClicked: {
            editTransferDialog.targetEmpId = root.targetEmpId
            editTransferDialog.recordId = 0
            editTransferDialog.dateInput = new Date().toISOString().split('T')[0]
            editTransferDialog.groupId = 0
            
            if (ApplicationWindow.window) {
                editTransferDialog.x = (ApplicationWindow.window.width - editTransferDialog.width) / 2
                editTransferDialog.y = (ApplicationWindow.window.height - editTransferDialog.height) / 2
            }
            editTransferDialog.open()
        }
    }

    Item { width: parent.width; implicitHeight: AppTheme.spaceS; visible: backend.employeeTransferHistory.length > 0 }

    Column {
        width: parent.width
        spacing: AppTheme.spaceM
        visible: backend.employeeTransferHistory.length > 0

        Repeater {
            model: backend.employeeTransferHistory

            Rectangle {
                width: parent.width
                height: 50
                radius: AppTheme.radiusMedium
                color: AppTheme.bgSurface 
                border.color: AppTheme.borderDivider
                border.width: 1
                
                layer.enabled: true
                layer.effect: DropShadow { transparentBorder: true; color: AppTheme.shadowColor; radius: AppTheme.shadowL1Blur; verticalOffset: AppTheme.shadowL1Y; samples: 9 }

                RowLayout {
                    anchors.fill: parent
                    anchors.margins: AppTheme.spaceM
                    anchors.rightMargin: 80 
                    spacing: AppTheme.spaceL
                    
                    Text { 
                        text: modelData.date
                        color: AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        Layout.preferredWidth: 80 
                    }
                    
                    Text { 
                        text: modelData.group_name
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        font.weight: AppTheme.weightBold
                        Layout.fillWidth: true
                        elide: Text.ElideRight 
                    }
                }

                Row {
                    anchors.right: parent.right
                    anchors.rightMargin: AppTheme.spaceS
                    anchors.verticalCenter: parent.verticalCenter
                    spacing: AppTheme.spaceXS
                    
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: editHov.pressed ? AppTheme.statePress : (editHov.containsMouse ? AppTheme.stateHover : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        
                        IconImage { anchors.centerIn: parent; source: "../icons/edit.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: editHov.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary }
                        MouseArea { 
                            id: editHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                editTransferDialog.targetEmpId = root.targetEmpId; editTransferDialog.recordId = modelData.id; editTransferDialog.dateInput = modelData.raw_date; editTransferDialog.groupId = modelData.group_id
                                if (ApplicationWindow.window) { editTransferDialog.x = (ApplicationWindow.window.width - editTransferDialog.width) / 2; editTransferDialog.y = (ApplicationWindow.window.height - editTransferDialog.height) / 2 }
                                editTransferDialog.open()
                            }
                        }
                    }
                    
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall
                        color: delHov.pressed ? AppTheme.statePress : (delHov.containsMouse ? AppTheme.bgDangerSoft : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        
                        IconImage { anchors.centerIn: parent; source: "../icons/trash.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: delHov.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary }
                        MouseArea { id: delHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: backend.deleteTransferRecord(root.targetEmpId, modelData.id) }
                    }
                }
            }
        }
    }
    
    Item {
        width: parent.width; implicitHeight: 200 
        visible: backend.employeeTransferHistory.length === 0
        Text { text: "История переводов пуста."; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; anchors.centerIn: parent }
    }

    AppDialog {
        id: editTransferDialog
        width: 300
        property int targetEmpId: 0; property int recordId: 0; property string dateInput: ""; property int groupId: 0
        title: recordId === 0 ? "Новая запись" : "Редактирование"
        acceptText: "Сохранить"

        onOpened: {
            for (let i = 0; i < backend.groupList.length; i++) {
                if (backend.groupList[i].id === groupId) { transferGroupCombo.currentIndex = i; break }
            }
        }

        AppDateField { id: editTransferDate; width: parent.width; label: "Дата:"; selectedDate: editTransferDialog.dateInput }
        AppComboBox { id: transferGroupCombo; width: parent.width; label: "Группа:"; model: backend.groupList; textRole: "name"; valueRole: "id" }

        onAccepted: { backend.saveTransferRecord(editTransferDialog.targetEmpId, editTransferDialog.recordId, editTransferDate.selectedDate, transferGroupCombo.currentValue); editTransferDialog.close() }
    }
}