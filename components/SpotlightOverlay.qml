import QtQuick
import Qt5Compat.GraphicalEffects

Item {
    id: root
    anchors.fill: parent
    z: AppTheme.zModal - 1
    
    // Входящие данные
    property Item targetItem: null
    property Item backgroundSource: null
    property bool isActive: false

    // Анимация появления
    visible: opacity > 0
    opacity: isActive ? 1.0 : 0.0
    Behavior on opacity { 
        NumberAnimation { duration: AppTheme.speedStandard; easing.type: AppTheme.speedStandard } 
    }

    // 1. Сильное размытие заднего фона
    FastBlur {
        anchors.fill: parent
        source: root.backgroundSource
        radius: 48
        transparentBorder: false
    }

    // 2. Темная пленка
    Rectangle {
        anchors.fill: parent
        color: AppTheme.isDark ? Qt.rgba(0, 0, 0, 0) : Qt.rgba(0, 0, 0, 0.1)
    }

    // 3. Контейнер для голограммы
    Item {
        id: cloneWrapper
        
        ShaderEffectSource {
            anchors.fill: parent
            sourceItem: root.targetItem
            live: true 
        }
    }

    // 4. Безопасный трекер координат (работает только когда эффект включен)
    Timer {
        interval: 16
        running: root.isActive && root.targetItem !== null
        repeat: true
        onTriggered: {
            if (root.targetItem) {
                // Высчитываем координаты относительно этого оверлея
                let pt = root.targetItem.mapToItem(root, 0, 0);
                cloneWrapper.x = pt.x;
                cloneWrapper.y = pt.y;
                cloneWrapper.width = root.targetItem.width;
                cloneWrapper.height = root.targetItem.height;
            }
        }
    }
}