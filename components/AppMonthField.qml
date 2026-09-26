import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

TextField {
    id: root

    property string label: ""
    property bool isRequired: false 
    property string selectedMonth: "" 
    property int currentYear: new Date().getFullYear()
    
    property color cutoutColor: AppTheme.bgModal
    property bool isFloated: root.text.length > 0 || root.activeFocus || monthDialog.opened

    implicitHeight: 44 
    Layout.fillWidth: true
    
    leftPadding: AppTheme.spaceM
    rightPadding: 36
    verticalAlignment: TextInput.AlignVCenter
    
    color: root.enabled ? AppTheme.textPrimary : AppTheme.textDisabled
    font.family: AppTheme.fontFamily
    font.pixelSize: AppTheme.sizeBody
    
    placeholderText: "ММ.ГГГГ"
    placeholderTextColor: {
        if ((root.activeFocus || monthDialog.opened) && floatingLabel.y < 0) {
            return AppTheme.textTertiary; 
        } else {
            return Qt.rgba(AppTheme.textTertiary.r, AppTheme.textTertiary.g, AppTheme.textTertiary.b, 0.0);
        }
    }
    Behavior on placeholderTextColor { ColorAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutQuad } }
    
    focusPolicy: Qt.StrongFocus
    cursorDelegate: AppCursorDelegate {}

    onActiveFocusChanged: { if (activeFocus) Qt.callLater(function() { root.selectAll() }) }

    onTextEdited: {
        let raw = text.replace(/[^0-9]/g, '')
        if (raw.length > 6) raw = raw.substring(0, 6)
        let formatted = ''
        if (raw.length > 0) formatted += raw.substring(0, 2)
        if (raw.length >= 3) formatted += '.' + raw.substring(2, 6)
        if (text !== formatted) { text = formatted; cursorPosition = formatted.length }
    }

    onEditingFinished: root.parseManualInput(root.text)

    // ==========================================
    // РАМКА
    // ==========================================
    background: Rectangle {
        color: "transparent"
        radius: AppTheme.radiusMedium
        
        border.color: !root.enabled ? AppTheme.borderDisabled :
                      (root.activeFocus || monthDialog.opened ? AppTheme.borderFocus : 
                      (root.hovered ? AppTheme.textSecondary : AppTheme.borderInput))
        
        border.width: (root.activeFocus || monthDialog.opened) ? AppTheme.focusWidth : 1
        Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ИКОНКА КАЛЕНДАРЯ
    Item {
        width: 36; height: 36
        anchors.right: parent.right
        anchors.verticalCenter: parent.verticalCenter
        
        IconImage { 
            anchors.centerIn: parent
            source: "../icons/calendar.svg"
            width: AppTheme.iconMedium; height: AppTheme.iconMedium
            color: monthDialog.opened ? AppTheme.accentBrand : AppTheme.textSecondary
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } } 
        }
        MouseArea {
            anchors.fill: parent; cursorShape: Qt.PointingHandCursor
            onClicked: { 
                if (monthDialog.opened) monthDialog.close(); 
                else monthDialog.showAt(root, root.width / 2, root.height) 
            }
        }
    }

    // ==========================================
    // ИДЕАЛЬНЫЙ ЛАСТИК И ЛЕЙБЛ
    // ==========================================
    Rectangle {
        color: root.cutoutColor
        x: floatingLabel.x - 4
        y: -2 
        height: 4 
        width: (floatingLabel.width * floatingLabel.scale) + 8
        opacity: root.isFloated ? 1.0 : 0.0
        Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
    }

    Row {
        id: floatingLabel
        x: AppTheme.spaceS
        y: root.isFloated ? -(height * 0.75) / 2 : (root.height - height) / 2
        
        scale: root.isFloated ? 0.75 : 1.0
        transformOrigin: Item.TopLeft
        
        Behavior on y { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }
        Behavior on scale { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }

        spacing: AppTheme.spaceMicro

        Text {
            text: root.label
            color: !root.enabled ? AppTheme.textDisabled : 
                   ((root.activeFocus || monthDialog.opened) ? AppTheme.accentBrand : AppTheme.textSecondary)
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
            font.weight: root.isFloated ? AppTheme.weightMedium : AppTheme.weightRegular
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        }
        
        Text {
            visible: root.isRequired
            text: "*"
            color: root.enabled ? AppTheme.accentDanger : AppTheme.textDisabled 
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
        }
    }

    // ==========================================
    // 3. ВЫПАДАЮЩЕЕ ОКНО (Выбор месяца)
    // ==========================================
    AppDialog {
        id: monthDialog
        parent: Overlay.overlay 
        width: 320 
        title: ""
        
        showFooter: false 

        onOpened: {
            if (root.selectedMonth !== "") {
                let parts = root.selectedMonth.split("-")
                if (parts.length === 2) root.currentYear = parseInt(parts[0])
            } else { 
                root.currentYear = new Date().getFullYear() 
            }
        }

        RowLayout {
            width: parent.width
            
            Rectangle {
                width: 32; height: 32; radius: AppTheme.radiusSmall
                color: prevHover.pressed ? AppTheme.statePress : (prevHover.containsMouse ? AppTheme.stateHover : "transparent")
                Text { anchors.centerIn: parent; text: "‹"; color: AppTheme.textSecondary; font.pixelSize: AppTheme.sizeH2 }
                MouseArea { id: prevHover; anchors.fill: parent; hoverEnabled: true; onClicked: root.currentYear-- }
            }
            
            Text {
                Layout.fillWidth: true; horizontalAlignment: Text.AlignHCenter
                text: root.currentYear.toString()
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH2
                font.weight: AppTheme.weightBold
            }
            
            Rectangle {
                width: 32; height: 32; radius: AppTheme.radiusSmall
                color: nextHover.pressed ? AppTheme.statePress : (nextHover.containsMouse ? AppTheme.stateHover : "transparent")
                Text { anchors.centerIn: parent; text: "›"; color: AppTheme.textSecondary; font.pixelSize: AppTheme.sizeH2 }
                MouseArea { id: nextHover; anchors.fill: parent; hoverEnabled: true; onClicked: root.currentYear++ }
            }
        }

        GridLayout {
            width: parent.width
            columns: 4
            rowSpacing: AppTheme.spaceS
            columnSpacing: AppTheme.spaceS
            
            Repeater {
                model: ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
                
                Rectangle {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 46
                    radius: AppTheme.radiusMedium
                    
                    property string thisMonthVal: root.currentYear + "-" + ("0" + (index + 1)).slice(-2)
                    property bool isSelected: root.selectedMonth === thisMonthVal
                    
                    color: isSelected ? AppTheme.accentBrand : 
                           (monthHover.pressed ? AppTheme.statePress : 
                           (monthHover.containsMouse ? AppTheme.stateHover : AppTheme.bgBase)) 
                           
                    Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
                    
                    Text { 
                        anchors.centerIn: parent
                        text: modelData
                        color: isSelected ? AppTheme.textOnAccent : AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        font.weight: isSelected ? AppTheme.weightBold : AppTheme.weightMedium 
                    }

                    MouseArea {
                        id: monthHover
                        anchors.fill: parent
                        hoverEnabled: true
                        cursorShape: Qt.PointingHandCursor
                        
                        onClicked: {
                            root.selectedMonth = parent.thisMonthVal
                            monthDialog.close()
                        }
                    }
                }
            }
        }
    }

    onSelectedMonthChanged: {
        if (selectedMonth === "") { root.text = ""; return }
        let parts = selectedMonth.split("-")
        if (parts.length === 2) root.text = parts[1] + "." + parts[0]
    }

    function parseManualInput(txt) {
        let parts = txt.split(".")
        if (parts.length === 2 && parts[1].length === 4) {
            let m = parseInt(parts[0]); let y = parseInt(parts[1])
            if (m >= 1 && m <= 12) { 
                let monthStr = ("0" + m).slice(-2); 
                root.selectedMonth = y + "-" + monthStr; 
                return 
            }
        }
        root.selectedMonthChanged() 
    }
}