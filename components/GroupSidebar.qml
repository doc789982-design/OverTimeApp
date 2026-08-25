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
        height: 60 
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
        
        bottomMargin: 140 
        clip: true
        model: backend.groupList
        
        delegate: Item {
            id: grpDelegateItem
            width: ListView.view.width
            height: 56 + grpGap.height 

            DropArea {
                id: groupDropArea
                anchors.fill: parent
                onDropped: (drop) => {
                    if (drop.source && drop.source.empId !== undefined) {
                        if (drop.source.isShiftPressed) {
                            transferDialog.targetEmpId = drop.source.empId
                            transferDialog.targetGroupId = modelData.id
                            transferDialog.showAt(groupDropArea, drop.x, drop.y)
                        } else {
                            backend.moveEmployeeToGroup(drop.source.empId, modelData.id)
                        }
                    } else if (drop.source && drop.source.groupId !== undefined) {
                        backend.reorderGroups(drop.source.groupId, modelData.id)
                    }
                    drop.accept()
                }
            }

            Column {
                anchors.fill: parent

                // ==========================================
                // ЗАЗОР (При перетаскивании групп)
                // ==========================================
                Rectangle {
                    id: grpGap
                    width: parent.width
                    property bool isGroupDrag: groupDropArea.drag.source && groupDropArea.drag.source.groupId !== undefined
                    property bool showGap: groupDropArea.containsDrag && isGroupDrag && groupDropArea.drag.source.groupId !== modelData.id && modelData.id !== 0
                    
                    height: showGap ? AppTheme.spaceM : 0
                    color: "transparent"
                    clip: true
                    Behavior on height { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }

                    Rectangle { 
                        anchors.centerIn: parent
                        width: 32
                        height: 4
                        radius: AppTheme.radiusPill
                        color: AppTheme.accentBrand 
                    }
                }

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
                        cursorShape: Qt.PointingHandCursor
                        
                        drag.target: modelData.id !== 0 ? dragGroupProxy : null
                        
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
                        text: modelData.name // Здесь лежит ПОЛНОЕ имя группы
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