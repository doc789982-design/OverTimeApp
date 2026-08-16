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

    // ═══════════════════════════════════════════════════════════
    // ОПТИМИЗАЦИЯ ДЛЯ СЛАБЫХ ВИДЕОКАРТ:
    // Раньше размытие пересчитывалось КАЖДЫЙ КАДР (live: true),
    // т.е. видеокарта размывала весь экран 60 раз в секунду.
    // Теперь мы делаем "фотографию" фона ОДИН РАЗ в момент
    // открытия диалога и размываем только её. Фон за диалогом
    // всё равно неподвижен — глазом разницы нет, а нагрузка
    // падает в десятки раз.
    // ═══════════════════════════════════════════════════════════

    // 0. Замороженный снимок фона (обновляется только при открытии)
    ShaderEffectSource {
        id: frozenBackground
        anchors.fill: parent
        sourceItem: root.backgroundSource
        live: false          // НЕ обновлять каждый кадр!
        visible: false       // сам снимок не показываем, он нужен размытию
    }

    onIsActiveChanged: {
        if (isActive) {
            // Обновляем "фотографию" фона ровно один раз при открытии
            frozenBackground.scheduleUpdate()
        }
    }

    // 1. Сильное размытие снимка (вычисляется один раз, не каждый кадр)
    FastBlur {
        anchors.fill: parent
        source: frozenBackground
        radius: 48
        transparentBorder: false
        // Кэшируем результат размытия как обычную картинку:
        // пока диалог открыт, видеокарта просто показывает готовый кадр
        layer.enabled: true
        layer.smooth: true
    }

    // 2. Темная пленка
    Rectangle {
        anchors.fill: parent
        color: AppTheme.isDark ? Qt.rgba(0, 0, 0, 0) : Qt.rgba(0, 0, 0, 0.1)
    }

    // 3. Контейнер для голограммы (подсвеченная ячейка поверх размытия)
    Item {
        id: cloneWrapper
        
        ShaderEffectSource {
            anchors.fill: parent
            sourceItem: root.targetItem
            // Живая копия только пока эффект реально виден
            live: root.isActive
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
