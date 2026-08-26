import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects
import "."

Rectangle {
    id: root
    color: AppTheme.bgPanel
    
    property Item workspace: null

    function blurSearch() {
        searchInput.focus = false
    }

    function clearSearch() {
        searchInput.text = ""
        searchInput.forceActiveFocus()
    }

    function dismissSearchIfOutside(srcItem, x, y) {
        if (!searchFieldBox)
            return
        var p = searchFieldBox.mapFromItem(srcItem, x, y)
        if (p.x >= 0 && p.y >= 0 && p.x <= searchFieldBox.width && p.y <= searchFieldBox.height)
            return
        blurSearch()
    }

    Shortcut {
        sequence: "Escape"
        enabled: searchInput.activeFocus
        onActivated: root.blurSearch()
    }
    
    Rectangle { 
        width: 1
        color: AppTheme.borderDivider
        anchors.right: parent.right
        anchors.top: parent.top
        anchors.bottom: parent.bottom 
        z: AppTheme.zContent
    }

    // ==========================================
    // 1. ШАПКА
    // ==========================================
    Rectangle {
        id: headerArea
        height: AppTheme.barHeight 
        anchors.top: parent.top; anchors.left: parent.left; anchors.right: parent.right
        color: "transparent"

        Rectangle { anchors.bottom: parent.bottom; width: parent.width; height: 1; color: AppTheme.borderDivider }

        RowLayout {
            anchors.fill: parent; anchors.margins: AppTheme.spaceS; spacing: AppTheme.spaceM

            Rectangle {
                id: searchFieldBox
                Layout.fillWidth: true; Layout.fillHeight: true
                radius: AppTheme.radiusMedium
                color: searchInput.activeFocus ? AppTheme.bgElevated : AppTheme.bgBase
                border.color: searchInput.activeFocus ? AppTheme.borderFocus : AppTheme.borderDivider
                border.width: searchInput.activeFocus ? AppTheme.focusWidth : 1
                Behavior on color { ColorAnimation { duration: AppTheme.durNormal; easing.type: AppTheme.easeColor } }
                Behavior on border.color { ColorAnimation { duration: AppTheme.durNormal; easing.type: AppTheme.easeColor } }
                Behavior on border.width { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }

                Rectangle {
                    anchors.fill: parent
                    anchors.margins: -3
                    radius: parent.radius + 2
                    color: "transparent"
                    border.color: AppTheme.accentBrand
                    border.width: 2
                    opacity: searchInput.activeFocus ? 0.22 : 0
                    z: -1
                    Behavior on opacity { NumberAnimation { duration: AppTheme.durNormal; easing.type: AppTheme.easeStandard } }
                }

                RowLayout {
                    anchors.fill: parent
                    anchors.leftMargin: AppTheme.spaceM
                    anchors.rightMargin: AppTheme.spaceXS
                    spacing: AppTheme.spaceS

                    IconImage {
                        source: "../icons/search.svg"
                        width: AppTheme.iconMedium
                        height: AppTheme.iconMedium
                        color: searchInput.activeFocus ? AppTheme.accentBrand : AppTheme.textSecondary
                        Behavior on color { ColorAnimation { duration: AppTheme.durNormal } }
                    }
                    TextInput {
                        id: searchInput
                        Layout.fillWidth: true
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        verticalAlignment: TextInput.AlignVCenter
                        selectByMouse: true
                        Text {
                            text: "Поиск сотрудника"
                            color: AppTheme.textTertiary
                            visible: !parent.text && !parent.activeFocus
                            anchors.verticalCenter: parent.verticalCenter
                        }
                        onTextChanged: backend.setSearchText(text)
                    }
                    Item {
                        Layout.preferredWidth: searchInput.text.length > 0 ? 22 : 0
                        Layout.preferredHeight: 22
                        visible: Layout.preferredWidth > 0
                        clip: true

                        Behavior on Layout.preferredWidth {
                            NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard }
                        }

                        Rectangle {
                            anchors.fill: parent
                            radius: AppTheme.radiusPill
                            color: clearSearchMouse.containsMouse ? AppTheme.stateHover : "transparent"
                            IconImage {
                                anchors.centerIn: parent
                                source: "../icons/close.svg"
                                width: AppTheme.iconSmall
                                height: AppTheme.iconSmall
                                color: clearSearchMouse.containsMouse ? AppTheme.textPrimary : AppTheme.textTertiary
                            }
                            MouseArea {
                                id: clearSearchMouse
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onClicked: root.clearSearch()
                            }
                        }
                    }
                }
            }

            Item {
                Layout.preferredWidth: 36; Layout.preferredHeight: 36
                HoverHandler { id: filterHover }
                AnimatedFilterButton { id: activeFilterBtn; anchors.fill: parent; bgRadius: AppTheme.radiusMedium; onClicked: { backend.setActiveOnly(isActiveOnly) } }
                AppToolTip { anchors.horizontalCenter: parent.horizontalCenter; anchors.top: parent.bottom; anchors.topMargin: AppTheme.spaceXXS; dropDown: true; text: activeFilterBtn.isActiveOnly ? "Показать всех" : "Только активные"; isVisible: filterHover.hovered }
            }
        }
    }

    // ==========================================
    // 2. СПИСОК СОТРУДНИКОВ
    // ==========================================
    ListView {
        id: empList
        anchors.top: headerArea.bottom; anchors.bottom: parent.bottom; anchors.left: parent.left; anchors.right: parent.right
        bottomMargin: 140 + backend.updateChromeExtra; clip: true; model: backend.employeeList
        displaced: Transition { NumberAnimation { properties: "y"; duration: 180; easing.type: Easing.OutCubic } }
        move: Transition { NumberAnimation { properties: "y"; duration: 180; easing.type: Easing.OutCubic } }
        
        delegate: Item {
            id: empDelegateItem
            width: ListView.view.width
            readonly property int empCardHeight: Math.max(AppTheme.rowHeight, empTextCol.implicitHeight + AppTheme.spaceM)
            height: modelData.is_header ? 40 : empCardHeight

            DropArea {
                id: empDropArea
                anchors.fill: parent
                keys: ["employee"]
                enabled: !modelData.is_header
                property bool insertAfter: false
                onPositionChanged: (drag) => { insertAfter = drag.y > height * 0.5 }
                onDropped: (drop) => {
                    if (drop.source && drop.source.empId !== undefined && drop.source.empId !== modelData.id) {
                        backend.reorderEmployees(drop.source.empId, modelData.id, insertAfter)
                        drop.accept()
                    }
                }
            }

            Rectangle {
                visible: empDropArea.containsDrag && empDropArea.drag.source && empDropArea.drag.source.empId !== undefined && empDropArea.drag.source.empId !== modelData.id && !modelData.is_header
                height: 2
                radius: 1
                color: AppTheme.accentBrand
                anchors.left: parent.left
                anchors.right: parent.right
                anchors.leftMargin: AppTheme.spaceM
                anchors.rightMargin: AppTheme.spaceM
                y: empDropArea.insertAfter ? parent.height - 2 : 0
                z: 20
            }

            Column {
                anchors.fill: parent

                // ЗАГОЛОВОК ГРУППЫ
                Rectangle {
                    visible: modelData.is_header
                    width: parent.width; height: modelData.is_header ? 40 : 0; color: "transparent"
                    Text { anchors.bottom: parent.bottom; anchors.bottomMargin: AppTheme.spaceS; anchors.left: parent.left; anchors.leftMargin: AppTheme.spaceM; text: modelData.name.toUpperCase(); color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
                }

                // ==========================================
                // КАРТОЧКА СОТРУДНИКА (Material Design 3)
                // ==========================================
                Item {
                    id: cardContainer
                    visible: !modelData.is_header
                    width: parent.width
                    height: modelData.is_header ? 0 : empDelegateItem.empCardHeight
                    clip: true
                    opacity: empMouseArea.drag.active ? 0.35 : 1
                    Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }

                    property bool isSelected: backend.selectedEmployeeId === modelData.id

                    // 1. ФОН (Заливка в стиле MD3)
                    Rectangle {
                        id: empCardBg
                        // Отступы по краям для эффекта "плавающей карточки"
                        anchors.fill: parent
                        anchors.leftMargin: AppTheme.spaceXS
                        anchors.rightMargin: AppTheme.spaceXS
                        anchors.topMargin: 2
                        anchors.bottomMargin: 2
                        
                        radius: AppTheme.radiusLarge // MD3 любит большие скругления (12-16px)
                        
                        // Цвет: Активный -> Синий мягкий, Наведение -> Серый мягкий, Покой -> Прозрачный
                        color: cardContainer.isSelected ? AppTheme.bgBrandSoft : 
                               (empMouseArea.containsMouse ? AppTheme.stateHover : "transparent")

                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                        // Анимация вжатия при клике
                        scale: empMouseArea.pressed ? 0.98 : 1.0
                        Behavior on scale { NumberAnimation { duration: 100 } }
                    }

                    // 2. КОНТЕНТ
                    Column {
                        id: empTextCol
                        anchors.verticalCenter: parent.verticalCenter
                        anchors.left: parent.left
                        anchors.leftMargin: AppTheme.spaceL
                        anchors.right: parent.right
                        anchors.rightMargin: AppTheme.spaceL + AppTheme.spaceS
                        spacing: 2

                        Text {
                            id: empName
                            width: parent.width
                            text: modelData.name
                            color: cardContainer.isSelected ? AppTheme.accentBrand : (modelData.is_active ? AppTheme.textPrimary : AppTheme.textDisabled)
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: cardContainer.isSelected ? AppTheme.weightBold : AppTheme.weightMedium
                            elide: Text.ElideRight
                            maximumLineCount: 1
                            lineHeight: 1.15
                            lineHeightMode: Text.ProportionalHeight
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        }

                        Text {
                            width: parent.width
                            text: modelData.subtitle
                            color: modelData.is_active ? AppTheme.textSecondary : AppTheme.textDisabled
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            wrapMode: Text.WordWrap
                            maximumLineCount: 2
                            elide: Text.ElideRight
                            lineHeight: 1.15
                            lineHeightMode: Text.ProportionalHeight
                        }
                    }
                    
                    // ИНДИКАТОР НОРМЫ — тонкая полоска справа, как термометр
                    Item {
                        visible: modelData.has_overtime
                        width: 3
                        anchors.right: parent.right
                        anchors.rightMargin: AppTheme.spaceS
                        anchors.top: parent.top
                        anchors.topMargin: AppTheme.spaceS
                        anchors.bottom: parent.bottom
                        anchors.bottomMargin: AppTheme.spaceS

                        property real ratio: cardContainer.isSelected
                            ? backend.selectedEmployeeRatio
                            : (Math.max(0, modelData.shift_minutes) / Math.max(1, modelData.norm_minutes))

                        // Фоновая дорожка — цвет фона панели, почти невидима
                        Rectangle {
                            anchors.fill: parent
                            radius: 2
                            color: AppTheme.bgDisabled
                            opacity: 0.5
                        }

                        // Заполненная часть — снизу вверх
                        Rectangle {
                            width: parent.width
                            radius: 2
                            anchors.bottom: parent.bottom

                            property real clampedRatio: isNaN(parent.ratio) ? 0 : Math.min(1.0, parent.ratio)
                            height: clampedRatio * parent.height

                            // Цвет: если выбран — акцент бренда, если нет — просто borderDivider чуть ярче
                            color: cardContainer.isSelected
                                ? AppTheme.accentBrand
                                : AppTheme.borderInput

                            opacity: cardContainer.isSelected ? 0.9 : 0.7

                            Behavior on height {
                                NumberAnimation {
                                    duration: AppTheme.durStandard
                                    easing.type: AppTheme.easeStandard
                                }
                            }
                            Behavior on color {
                                ColorAnimation { duration: AppTheme.durStandard }
                            }
                        }

                        // Маленькая точка сверху когда норма выполнена — знак завершённости
                        Rectangle {
                            visible: parent.ratio >= 1.0
                            width: 5; height: 5; radius: 3
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.top: parent.top
                            anchors.topMargin: -1
                            color: cardContainer.isSelected
                                ? AppTheme.accentBrand
                                : AppTheme.borderInput
                            opacity: 0.9

                            Behavior on color {
                                ColorAnimation { duration: AppTheme.durStandard }
                            }
                        }
                    }

                    AppMenu {
                        id: empContextMenu
                        AppMenuItem { text: "Редактировать"; iconSource: "../icons/edit.svg"; onClicked: { empDialog.editId = modelData.id; empDialog.lastName = modelData.last_name; empDialog.firstName = modelData.first_name; empDialog.middleName = modelData.middle_name; empDialog.rank = modelData.rank; empDialog.position = modelData.position; empDialog.startMonth = modelData.start_month; empDialog.openHours = Math.floor(modelData.opening_minutes / 60).toString(); empDialog.openOvertime = Math.floor(modelData.opening_overtime / 60).toString(); empDialog.openDays = modelData.opening_days.toString(); empDialog.prevOpenHours = Math.floor(modelData.prev_opening_minutes / 60).toString(); empDialog.prevOpenOvertime = Math.floor(modelData.prev_opening_overtime / 60).toString(); empDialog.prevOpenDays = modelData.prev_opening_days.toString(); empDialog.showAt(cardContainer, cardContainer.width / 2, cardContainer.height / 2) } }
                        AppMenuItem { text: "История переводов"; iconSource: "../icons/clock.svg"; onClicked: { historyDialog.targetEmpId = modelData.id; historyDialog.targetEmpName = modelData.name; backend.loadTransferHistory(modelData.id); historyDialog.show() } }
                        AppMenuSeparator {}
                        AppMenuItem { text: "Переведен"; iconSource: "../icons/user_minus.svg"; onClicked: { endStatusDialog.targetEmpId = modelData.id; endStatusDialog.targetReason = "transfer"; endStatusDialog.showAt(cardContainer, cardContainer.width / 2, cardContainer.height / 2) } }
                        AppMenuItem { text: "Уволен"; iconSource: "../icons/user_minus.svg"; onClicked: { endStatusDialog.targetEmpId = modelData.id; endStatusDialog.targetReason = "dismissal"; endStatusDialog.showAt(cardContainer, cardContainer.width / 2, cardContainer.height / 2) } }
                        AppMenuItem { visible: !modelData.is_active; text: "Отменить статус"; iconSource: "../icons/x_circle.svg"; onClicked: backend.clearEmployeeEndDate(modelData.id) }
                        AppMenuSeparator {}
                        AppMenuItem {
                            text: "Удалить"
                            iconSource: "../icons/trash.svg"
                            isDanger: true
                            onClicked: {
                                mainWindow.askConfirm(
                                    "Удалить сотрудника?",
                                    "«" + modelData.name + "» будет удалён вместе со всеми дежурствами и компенсациями.\nЕсли передумаете — нажмите Ctrl+Z.",
                                    "Удалить",
                                    function() { backend.deleteEmployee(modelData.id) }
                                )
                            }
                        }
                    }

                    MouseArea { 
                        id: empMouseArea
                        anchors.fill: parent; hoverEnabled: true; acceptedButtons: Qt.LeftButton | Qt.RightButton
                        drag.target: dragProxy
                        drag.threshold: 8
                        onPositionChanged: (mouse) => { if (drag.active && root.workspace) { let pt = mapToItem(root.workspace, mouse.x, mouse.y); dragProxy.x = pt.x - (dragProxy.width / 2); dragProxy.y = pt.y - (dragProxy.height / 2) } }
                        onPressed: (mouse) => { if (mouse.button === Qt.LeftButton && root.workspace) { dragProxy.isShiftPressed = (mouse.modifiers & Qt.ShiftModifier) !== 0; dragProxy.parent = root.workspace; let pt = mapToItem(root.workspace, mouse.x, mouse.y); dragProxy.x = pt.x - (dragProxy.width / 2); dragProxy.y = pt.y - (dragProxy.height / 2) } }
                        onReleased: { dragProxy.Drag.drop(); dragProxy.parent = cardContainer; dragProxy.x = 0; dragProxy.y = 0 }
                        onClicked: (mouse) => { if (mouse.button === Qt.RightButton) empContextMenu.popup(); else backend.selectEmployee(modelData.id) }
                    }

                    AppToolTip { anchors.horizontalCenter: parent.horizontalCenter; anchors.bottom: parent.top; anchors.bottomMargin: AppTheme.spaceXXS; text: modelData.inactive_reason || ""; isVisible: !modelData.is_active && empMouseArea.containsMouse && text !== "" }

                    Rectangle {
                        id: dragProxy
                        width: empDelegateItem.width - 32; height: 56; color: AppTheme.bgElevated; radius: AppTheme.radiusLarge
                        border.color: AppTheme.accentBrand; border.width: 1; opacity: 0.95     
                        property int empId: modelData.id; property bool isShiftPressed: false
                        Drag.active: empMouseArea.drag.active; Drag.keys: ["employee"]; Drag.hotSpot.x: width / 2; Drag.hotSpot.y: height / 2; visible: empMouseArea.drag.active
                        AppShadow { level: 4 }
                        Text { anchors.left: parent.left; anchors.leftMargin: AppTheme.spaceM; anchors.verticalCenter: parent.verticalCenter; text: modelData.name; color: AppTheme.textPrimary; font.pixelSize: AppTheme.sizeBody; font.weight: AppTheme.weightBold }
                    }
                }
            }
        }
    }
}