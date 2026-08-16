import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

MenuItem {
    id: control
    
    property bool isDanger: false
    property string customColor: ""
    property string iconSource: ""

    default property alias customContent: extraArea.data

    implicitHeight: visible ? 36 : 0 
    
    // Если кнопка заблокирована - она полупрозрачная
    opacity: control.enabled ? 1.0 : AppTheme.alphaDisabled

    contentItem: Item {
        
        // ИКОНКА (Если есть)
        IconImage {
            id: leftIcon
            visible: control.iconSource !== ""
            source: control.iconSource
            width: AppTheme.iconSmall // Строго 12px для меню
            height: AppTheme.iconSmall
            
            color: control.customColor !== "" ? control.customColor : 
                  (control.isDanger ? AppTheme.accentDanger : AppTheme.textPrimary)
                  
            anchors.verticalCenter: parent.verticalCenter
            anchors.left: parent.left
            anchors.leftMargin: AppTheme.spaceM
        }

        // ТЕКСТ
        Text {
            text: control.text
            color: control.customColor !== "" ? control.customColor : 
                  (control.isDanger ? AppTheme.accentDanger : AppTheme.textPrimary)
                  
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody
            font.weight: control.isDanger ? AppTheme.weightMedium : AppTheme.weightRegular // Опасные действия полужирные
            
            anchors.verticalCenter: parent.verticalCenter
            // Динамический отступ текста, если есть иконка
            anchors.left: leftIcon.visible ? leftIcon.right : parent.left
            anchors.leftMargin: leftIcon.visible ? AppTheme.spaceS : AppTheme.spaceM
        }
        
        // ЭКСТРА КОНТЕНТ (Крестики удаления справа и т.д.)
        Item {
            id: extraArea
            anchors.right: parent.right
            anchors.top: parent.top
            anchors.bottom: parent.bottom
        }
    }

    // ==========================================
    // ФОН (Интерактивное состояние)
    // ==========================================
    background: Rectangle {
        // МАГИЯ: Мягкий фон при наведении
        color: control.hovered ? AppTheme.stateHover : "transparent"
        radius: AppTheme.radiusSmall
        anchors.margins: AppTheme.spaceXXS // Маленький отступ от краев самого меню (4px)
    }
}