import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects 
import "."

Rectangle {
    id: root
    color: AppTheme.bgPanel // Фон совпадает с главным окном
    
    property int activeGroupId: 0
    property Item workspace: null 

    // ==========================================
    // 1. ШАПКА (Кнопка Добавления)
    // ==========================================
    Rectangle {
        id: headerArea
        height: AppTheme.barHeight 
        anchors.top: parent.top
        anchors.left: parent.left
        anchors.right: parent.right
        color: "transparent"

        // Горизонтальный разделитель
        Rectangle { 
            anchors.bottom: parent.bottom
            width: parent.width
            height: 1
            color: AppTheme.borderDivider 
        }

        // Кнопка "+"
        Rectangle {
            id: addGroupBtn
            width: 36
            height: 36
            radius: AppTheme.radiusPill 
            anchors.centerIn: parent
            
            color: addGrpMouse.pressed ? AppTheme.statePress : (addGrpMouse.containsMouse ? AppTheme.stateHover : "transparent")
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
            
            Text { 
                anchors.centerIn: parent
                text: "+"
                color: AppTheme.textSecondary
                font.weight: AppTheme.weightMedium
                font.pixelSize: AppTheme.sizeH2 
            }
            
            MouseArea { 
                id: addGrpMouse
                anchors.fill: parent
                hoverEnabled: true
                cursorShape: Qt.PointingHandCursor
                onClicked: function(mouse) { addGroupDialog.showAt(parent, mouse.x, mouse.y) }
            }
        }
    }

    // ==========================================
    // 2. СПИСОК ГРУПП
    // ==========================================
    ListView {
        anchors.top: headerArea.bottom
        anchors.bottom: parent.bottom 
        anchors.left: parent.left
        anchors.right: parent.right
        anchors.topMargin: AppTheme.spaceM
        
        bottomMargin: 140 + backend.updateChromeExtra 
        clip: true
        model: backend.groupList
        
        displaced: Transition { NumberAnimation { properties: "y"; duration: 180; easing.type: Easing.OutCubic } }
        move: Transition { NumberAnimation { properties: "y"; duration: 180; easing.type: Easing.OutCubic } }

        delegate: Item {
            id: grpDelegateItem
            width: ListView.view.width
            height: 56

            DropArea {
                id: groupDropArea
                anchors.fill: parent
                keys: ["group", "employee"]
                property bool insertAfter: false
                readonly property bool isGroupDrag: drag.source && drag.source.groupId !== undefined
                readonly property bool isEmpDrag: drag.source && drag.source.empId !== undefined
                onPositionChanged: (drag) => { insertAfter = drag.y > height * 0.5 }
                onDropped: (drop) => {
                    if (drop.source && drop.source.empId !== undefined) {
                        if (drop.source.isShiftPressed) {
                            transferDialog.targetEmpId = drop.source.empId
                            transferDialog.targetGroupId = modelData.id
                            transferDialog.showAt(groupDropArea, drop.x, drop.y)
                        } else {
                            backend.moveEmployeeToGroup(drop.source.empId, modelData.id)
                        }
                    } else if (drop.source && drop.source.groupId !== undefined && modelData.id !== 0 && drop.source.groupId !== modelData.id) {
                        backend.reorderGroups(drop.source.groupId, modelData.id, insertAfter)
                    }
                    drop.accept()
                }
            }

            Rectangle {
                visible: groupDropArea.containsDrag && groupDropArea.isGroupDrag && groupDropArea.drag.source.groupId !== modelData.id && modelData.id !== 0
                width: 28
                height: 3
                radius: AppTheme.radiusPill
                color: AppTheme.accentBrand
                anchors.horizontalCenter: parent.horizontalCenter
                y: groupDropArea.insertAfter ? parent.height - 4 : 2
                z: 20
            }

            Column {
                anchors.fill: parent

                // ==========================================
                // КНОПКА ГРУППЫ
                // ==========================================
                Item {
                    width: 48
                    height: 48
                    anchors.horizontalCenter: parent.horizontalCenter
                    
                    property bool isSelected: root.activeGroupId === modelData.id
                    property bool isEmpDrag: groupDropArea.containsDrag && groupDropArea.drag.source && groupDropArea.drag.source.empId !== undefined
                    
                    // ВЫЗЫВАЕМ НАШ КОМПОНЕНТ
                    AppSelectionIndicator {
                        // Жестко прибиваем к левому краю боковой панели
                        x: -12
                        anchors.verticalCenter: parent.verticalCenter
                        
                        targetHeight: parent.height
                        isHovered: btnMouseArea.containsMouse
                        isSelected: parent.isSelected
                        isDragged: parent.isEmpDrag
                    }

                    // Текст (Все, Смена 1, и т.д.)
                    Text { 
                        anchors.centerIn: parent
                        text: modelData.icon
                        
                        // Цвет текста меняется в зависимости от того, наведена мышь или выбрана группа
                        color: parent.isSelected ? AppTheme.accentBrand : 
                               (btnMouseArea.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary)
                               
                        font.family: AppTheme.fontFamily
                        font.weight: (parent.isSelected || btnMouseArea.containsMouse) ? AppTheme.weightBold : AppTheme.weightMedium
                        font.pixelSize: AppTheme.sizeBody 
                        
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                    }

                    // Контекстное меню
                    AppMenu {
                        id: groupMenu
                        AppMenuItem {
                            visible: modelData.id !== 0
                            text: modelData.shifted_weekends ? "Обычные выходные" : "Смещённые выходные"
                            iconSource: "../icons/calendar.svg"
                            onClicked: backend.setGroupShiftedWeekends(modelData.id, !modelData.shifted_weekends)
                        }
                        AppMenuSeparator { visible: modelData.id !== 0 }
                        AppMenuItem {
                            visible: modelData.id !== 0 
                            text: "Удалить группу"
                            isDanger: true 
                            iconSource: "../icons/trash.svg"
                            onClicked: {
                                let empCount = backend.getGroupEmployeeCount(modelData.id)
                                let msg = empCount === 0
                                    ? "Группа «" + modelData.name + "» пуста — можно удалять.\nЕсли передумаете — нажмите Ctrl+Z."
                                    : "В группе «" + modelData.name + "» сейчас " + empCount + " чел.\nПосле удаления они останутся в списке «Без группы»."
                                mainWindow.askConfirm(
                                    "Удалить группу?",
                                    msg,
                                    "Удалить",
                                    function() { backend.deleteGroup(modelData.id) }
                                )
                            }
                        }
                    }

                    // Зона управления мышью
                    MouseArea { 
                        id: btnMouseArea
                        anchors.fill: parent
                        hoverEnabled: true
                        acceptedButtons: Qt.LeftButton | Qt.RightButton
                        cursorShape: drag.active ? Qt.ClosedHandCursor : Qt.PointingHandCursor
                        drag.target: modelData.id !== 0 ? dragGroupProxy : null
                        drag.threshold: 10
                        
                        onPositionChanged: (mouse) => {
                            if (drag.active && root.workspace) {
                                let pt = mapToItem(root.workspace, mouse.x, mouse.y)
                                dragGroupProxy.x = pt.x - (dragGroupProxy.width / 2); dragGroupProxy.y = pt.y - (dragGroupProxy.height / 2)
                            }
                        }
                        onPressed: (mouse) => {
                            if (mouse.button === Qt.LeftButton && modelData.id !== 0 && root.workspace) {
                                dragGroupProxy.parent = root.workspace
                                let pt = mapToItem(root.workspace, mouse.x, mouse.y)
                                dragGroupProxy.x = pt.x - (dragGroupProxy.width / 2); dragGroupProxy.y = pt.y - (dragGroupProxy.height / 2)
                            }
                        }
                        onReleased: {
                            if (modelData.id !== 0) { dragGroupProxy.Drag.drop(); dragGroupProxy.parent = parent; dragGroupProxy.x = 0; dragGroupProxy.y = 0 }
                        }
                        onClicked: (mouse) => {
                            if (mouse.button === Qt.RightButton && modelData.id !== 0) { groupMenu.popup() } 
                            else { root.activeGroupId = modelData.id; backend.setGroupFilter(modelData.id) }
                        }
                    }

                    AppToolTip {
                        anchors.left: parent.right
                        anchors.leftMargin: AppTheme.spaceXS
                        anchors.verticalCenter: parent.verticalCenter
                        text: modelData.name
                              + (modelData.shifted_weekends ? " · смещённые выходные" : "")
                              + (modelData.is_shift ? " · сменный график" : "")
                        isVisible: btnMouseArea.containsMouse && !btnMouseArea.drag.active
                    }

                    // ФАНТОМ ПРИ ПЕРЕТАСКИВАНИИ ГРУППЫ
                    Rectangle {
                        id: dragGroupProxy
                        width: 48; height: 48; radius: AppTheme.radiusPill
                        
                        color: AppTheme.accentBrand
                        opacity: 0.9
                        
                        property int groupId: modelData.id 
                        Drag.active: btnMouseArea.drag.active
                        Drag.keys: ["group"]
                        Drag.hotSpot.x: 24; Drag.hotSpot.y: 24
                        visible: btnMouseArea.drag.active
                        
                        // Тень-картинка вместо вычисляемой (Level 3)
                        AppShadow { level: 3 }

                        Text { 
                            anchors.centerIn: parent
                            text: modelData.icon
                            color: AppTheme.textOnAccent
                            font.weight: AppTheme.weightBold
                            font.pixelSize: AppTheme.sizeH3 
                        }
                    }
                }
            }
        }
    }
}