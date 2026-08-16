import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

Item {
    id: root
    implicitWidth: 32
    implicitHeight: 32

    // Цвета строго как у кнопок печати и настроек
    property color iconColor: AppTheme.textSecondary
    property color iconHoverColor: AppTheme.textPrimary
    property int bgRadius: AppTheme.radiusMedium

    property bool isActiveOnly: true

    signal clicked()

    // 1. ФОН ПРИ НАВЕДЕНИИ
    Rectangle {
        anchors.fill: parent
        radius: root.bgRadius
        color: mouseArea.pressed ? AppTheme.statePress : (mouseArea.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // 2. КОНТЕЙНЕР ДЛЯ ИКОНОК
    Item {
        width: AppTheme.iconLarge
        height: AppTheme.iconLarge
        anchors.centerIn: parent

        // --- ИКОНКА 2: ВСЕ СОТРУДНИКИ (people_alt) ---
        IconImage {
            id: iconMultiple
            anchors.fill: parent
            source: "../icons/people_alt.svg"
            
            color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

            // Логика видимости
            opacity: root.isActiveOnly ? 0.0 : 1.0
            scale: root.isActiveOnly ? 0.5 : 1.0 // Слегка уменьшается при исчезновении
            
            Behavior on opacity { NumberAnimation { duration: 250; easing.type: Easing.InOutQuad } }
            Behavior on scale { NumberAnimation { duration: 250; easing.type: Easing.OutBack } }
        }

        // --- ИКОНКА 1: ТОЛЬКО АКТИВНЫЕ (person) ---
        IconImage {
            id: iconSingle
            anchors.fill: parent
            source: "../icons/person.svg"
            
            color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

            // Логика видимости
            opacity: root.isActiveOnly ? 1.0 : 0.0
            scale: root.isActiveOnly ? 1.0 : 0.5 // Слегка уменьшается при исчезновении
            
            Behavior on opacity { NumberAnimation { duration: 250; easing.type: Easing.InOutQuad } }
            Behavior on scale { NumberAnimation { duration: 250; easing.type: Easing.OutBack } }
        }
    }

    // 3. ЗОНА КЛИКА
    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        
        onClicked: {
            root.isActiveOnly = !root.isActiveOnly;
            root.clicked();
        }
    }
}