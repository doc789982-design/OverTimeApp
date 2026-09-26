import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import Qt5Compat.GraphicalEffects
import QtQuick.Controls.impl

ComboBox {
    id: control
    Layout.fillWidth: true
    implicitHeight: 44 

    property string label: ""
    property bool isRequired: false
    property color cutoutColor: AppTheme.bgModal
    
    // Считается активным, если есть текст, открыто меню или есть фокус
    property bool isFloated: control.displayText.length > 0 || control.activeFocus || control.popup.visible

    leftPadding: AppTheme.spaceM
    rightPadding: AppTheme.spaceXL 

    focusPolicy: Qt.StrongFocus

    // ==========================================
    // ТЕКСТ ВНУТРИ
    // ==========================================
    contentItem: Text {
        text: control.displayText 
        color: control.enabled ? AppTheme.textPrimary : AppTheme.textDisabled
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        verticalAlignment: Text.AlignVCenter
        
        elide: Text.ElideRight 
    }

    // ==========================================
    // РАМКА
    // ==========================================
    background: Rectangle {
        color: "transparent"
        radius: AppTheme.radiusMedium
        
        border.color: !control.enabled ? AppTheme.borderDisabled :
                      (control.activeFocus || control.popup.visible ? AppTheme.borderFocus : 
                      (control.hovered ? AppTheme.textSecondary : AppTheme.borderInput))
        border.width: (control.activeFocus || control.popup.visible) ? AppTheme.focusWidth : 1
        
        Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ==========================================
    // ИДЕАЛЬНЫЙ ЛАСТИК И ЛЕЙБЛ
    // ==========================================
    Rectangle {
        color: control.cutoutColor
        x: floatingLabel.x - 4
        y: -2 
        height: 4 
        width: (floatingLabel.width * floatingLabel.scale) + 8
        opacity: control.isFloated ? 1.0 : 0.0
        Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
    }

    Row {
        id: floatingLabel
        x: AppTheme.spaceM
        y: control.isFloated ? -(height * 0.75) / 2 : (control.height - height) / 2
        
        scale: control.isFloated ? 0.75 : 1.0
        transformOrigin: Item.TopLeft
        
        Behavior on y { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }
        Behavior on scale { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }

        spacing: AppTheme.spaceMicro

        Text {
            text: control.label
            color: !control.enabled ? AppTheme.textDisabled : 
                   ((control.activeFocus || control.popup.visible) ? AppTheme.accentBrand : AppTheme.textSecondary)
                   
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
            font.weight: control.isFloated ? AppTheme.weightMedium : AppTheme.weightRegular
            
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        }
        
        Text {
            visible: control.isRequired
            text: "*"
            color: control.enabled ? AppTheme.accentDanger : AppTheme.textDisabled 
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
        }
    }

    // ==========================================
    // СТРЕЛОЧКА И ВЫПАДАЮЩЕЕ МЕНЮ
    // ==========================================
    indicator: Item {
        x: control.width - width - AppTheme.spaceS
        y: (control.availableHeight - height) / 2
        width: AppTheme.iconLarge
        height: AppTheme.iconLarge
        
        IconImage {
            anchors.centerIn: parent
            source: "../icons/chevron_down.svg"
            width: AppTheme.iconSmall
            height: AppTheme.iconSmall
            color: control.enabled ? AppTheme.textSecondary : AppTheme.textDisabled
            rotation: control.popup.visible ? 180 : 0
            Behavior on rotation { NumberAnimation { duration: AppTheme.durNormal; easing.type: AppTheme.easeStandard } }
        }
    }

    popup: Popup {
        y: control.height + AppTheme.spaceXXS 
        width: control.width
        implicitHeight: contentItem.implicitHeight + (padding * 2)
        padding: AppTheme.spaceXXS 
        z: AppTheme.zDropdown 

        enter: Transition {
            ParallelAnimation {
                NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
                NumberAnimation { property: "y"; from: root.y - AppTheme.slideOffset; to: control.height + AppTheme.spaceXXS; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            }
        }
        exit: Transition {
            ParallelAnimation {
                NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
                NumberAnimation { property: "y"; from: control.height + AppTheme.spaceXXS; to: root.y - AppTheme.slideOffset; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
            }
        }

        background: Rectangle {
            color: AppTheme.bgElevated
            border.color: AppTheme.borderDivider
            border.width: 1
            radius: AppTheme.radiusMedium
            // Тень-картинка вместо вычисляемой (Level 2)
            AppShadow { level: 2 }
        }

        contentItem: ListView {
            clip: true
            implicitHeight: contentHeight
            model: control.delegateModel 
            currentIndex: control.highlightedIndex
            spacing: 2 

            delegate: ItemDelegate {
                width: ListView.view.width
                height: 36 
                property bool isSelected: control.currentIndex === index
                
                contentItem: Text {
                    text: {
                        if (control.textRole && modelData[control.textRole] !== undefined) return modelData[control.textRole]
                        if (modelData.text !== undefined) return modelData.text
                        return modelData
                    }
                    color: isSelected ? AppTheme.accentBrand : AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: isSelected ? AppTheme.weightBold : AppTheme.weightRegular 
                    verticalAlignment: Text.AlignVCenter
                    leftPadding: AppTheme.spaceS
                }
                
                background: Rectangle {
                    color: isSelected ? AppTheme.stateSelected : (parent.hovered || parent.highlighted ? AppTheme.stateHover : "transparent")
                    radius: AppTheme.radiusSmall
                }
            }
        }
    }
}