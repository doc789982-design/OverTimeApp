import QtQuick

// ============================================================
// РАДАР КНОПКИ ОБНОВЛЕНИЯ: точка в правом верхнем углу и исходящие
// от неё полупрозрачные кольца — «нажми на меня». Показывается,
// когда есть что скачать или установить. Два кольца расходятся
// в противофазе — сигнал «идёт волнами», как эхо-локатор.
// Цвет — акцентный (синий) в обоих состояниях: различает их стрелка
// (до загрузки — синяя, «можно установить» — зелёная).
// ============================================================
Item {
    id: root
    width: 22
    height: 22

    property bool active: false
    property color waveColor: AppTheme.accentBrand

    // Вся точка скрыта, когда сигнала нет (нет обновления / идёт загрузка)
    opacity: root.active ? 1.0 : 0.0
    Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }

    // Точка-источник сигнала
    Rectangle {
        id: dot
        anchors.centerIn: parent
        width: 6; height: 6; radius: 3
        color: root.waveColor
        // точка дышит в такт первому кольцу
        SequentialAnimation on scale {
            running: root.active; loops: Animation.Infinite
            NumberAnimation { to: 1.25; duration: 500; easing.type: Easing.OutQuad }
            NumberAnimation { to: 1.0; duration: 1100; easing.type: Easing.InQuad }
        }
    }

    // Кольцо сигнала: растёт от точки и тает
    Rectangle {
        id: wave1
        anchors.centerIn: parent
        width: 20; height: 20; radius: 10
        color: "transparent"
        border.color: root.waveColor
        border.width: 2
        opacity: 0
        SequentialAnimation on opacity {
            running: root.active; loops: Animation.Infinite
            NumberAnimation { to: 0.45; duration: 240; easing.type: Easing.OutQuad }
            NumberAnimation { to: 0.0; duration: 1360; easing.type: Easing.InQuad }
        }
        SequentialAnimation on scale {
            running: root.active; loops: Animation.Infinite
            NumberAnimation { from: 0.3; to: 1.0; duration: 1600; easing.type: Easing.OutQuad }
        }
    }

    // Второе кольцо — в противофазе (половина периода)
    Rectangle {
        id: wave2
        anchors.centerIn: parent
        width: 20; height: 20; radius: 10
        color: "transparent"
        border.color: root.waveColor
        border.width: 1.6
        opacity: 0
        SequentialAnimation on opacity {
            running: root.active; loops: Animation.Infinite
            PauseAnimation { duration: 800 }
            NumberAnimation { to: 0.35; duration: 240; easing.type: Easing.OutQuad }
            NumberAnimation { to: 0.0; duration: 1360; easing.type: Easing.InQuad }
        }
        SequentialAnimation on scale {
            running: root.active; loops: Animation.Infinite
            PauseAnimation { duration: 800 }
            NumberAnimation { from: 0.3; to: 1.0; duration: 1600; easing.type: Easing.OutQuad }
        }
    }
}
