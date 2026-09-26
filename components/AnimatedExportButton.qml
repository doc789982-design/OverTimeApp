import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

Item {
    id: root
    
    implicitWidth: 32
    implicitHeight: 32

    // Иконки всегда цвета Secondary, при наведении - Primary
    property color iconColor: AppTheme.textSecondary
    property color iconHoverColor: AppTheme.textPrimary
    property int bgRadius: AppTheme.radiusMedium

    signal clicked()

    // 1. ФОН И СОСТОЯНИЯ (Hover / Press)
    Rectangle {
        anchors.fill: parent
        radius: root.bgRadius
        color: mouseArea.pressed ? AppTheme.statePress : (mouseArea.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // 2. ИКОНКА 
    Item {
        width: AppTheme.iconLarge
        height: AppTheme.iconLarge
        anchors.centerIn: parent

        // Нижняя коробочка
        IconImage {
            id: boxLayer
            source: "../icons/export_box.svg"
            width: AppTheme.iconLarge; height: AppTheme.iconLarge
            color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        }

        // Стрелочка
        IconImage {
            id: arrowLayer
            source: "../icons/export_arrow.svg"
            width: AppTheme.iconLarge; height: AppTheme.iconLarge
            y: 0
            color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        }
    }

    // 3. АНИМАЦИЯ 
    SequentialAnimation {
        id: exportAnim
        
        ParallelAnimation {
            NumberAnimation { target: arrowLayer; property: "y"; to: 6; duration: AppTheme.durStandard; easing.type: AppTheme.easeExit }
            NumberAnimation { target: arrowLayer; property: "opacity"; to: 0.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeColor }
            
            SequentialAnimation {
                PauseAnimation { duration: 100 }
                NumberAnimation { target: boxLayer; property: "scale"; to: 0.85; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
                NumberAnimation { target: boxLayer; property: "scale"; to: 1.0; duration: AppTheme.durFast; easing.type: Easing.OutBounce }
            }
        }
        
        PropertyAction { target: arrowLayer; property: "y"; value: -10 }
        
        ParallelAnimation {
            NumberAnimation { target: arrowLayer; property: "opacity"; to: 1.0; duration: AppTheme.durFast }
            NumberAnimation { target: arrowLayer; property: "y"; to: 0; duration: AppTheme.durStandard; easing.type: Easing.OutBack }
        }
    }

    // 4. ЗОНА КЛИКА
    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor // Всегда добавляем "руку" для кнопок
        
        onEntered: { if (!exportAnim.running) exportAnim.start() }
        onClicked: { if (!exportAnim.running) exportAnim.start(); root.clicked() }
    }
}