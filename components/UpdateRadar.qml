import QtQuick

// ============================================================
// МАЯЧОК КНОПКИ ОБНОВЛЕНИЯ: точка в правом верхнем углу и одна
// спокойная исходящая волна — «нажми на меня». Показывается,
// когда есть что скачать или установить.
//
// Характер — сдержанный (сборка 310): волна одна вместо двух,
// период долгий, кольцо тонкое и тает мягко; точка дышит в том
// же ритме. Цвет — акцентный синий в обоих состояниях: их
// различает сам значок кнопки (синий — «скачать», зелёный —
// «установить»). Настройка «Меньше анимаций» оставляет только
// неподвижную точку.
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

    // Появление — с лёгким «всплытием» из уменьшенной точки
    scale: root.active ? 1.0 : 0.6
    Behavior on scale { NumberAnimation { duration: AppTheme.reducedMotion ? 0 : AppTheme.durSlow; easing.type: AppTheme.easeEnter } }

    // Волна: медленно расходится от точки и тает
    Rectangle {
        id: wave
        anchors.centerIn: parent
        width: 20; height: 20; radius: 10
        color: "transparent"
        border.color: root.waveColor
        border.width: 1.4
        opacity: 0
        SequentialAnimation on opacity {
            running: root.active && !AppTheme.reducedMotion; loops: Animation.Infinite
            NumberAnimation { to: 0.38; duration: 380; easing.type: Easing.OutQuad }
            NumberAnimation { to: 0.0; duration: 2020; easing.type: Easing.InQuad }
            PauseAnimation { duration: 500 }
        }
        SequentialAnimation on scale {
            running: root.active && !AppTheme.reducedMotion; loops: Animation.Infinite
            NumberAnimation { from: 0.3; to: 1.05; duration: 2400; easing.type: Easing.OutCubic }
            PauseAnimation { duration: 500 }
        }
    }

    // Точка-источник сигнала: дышит медленно и едва заметно
    Rectangle {
        id: dot
        anchors.centerIn: parent
        width: 6; height: 6; radius: 3
        color: root.waveColor
        SequentialAnimation on scale {
            running: root.active && !AppTheme.reducedMotion; loops: Animation.Infinite
            NumberAnimation { to: 1.18; duration: 1450; easing.type: Easing.InOutSine }
            NumberAnimation { to: 1.0; duration: 1450; easing.type: Easing.InOutSine }
        }
    }
}
