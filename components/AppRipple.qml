import QtQuick

// ============================================================
// RIPPLE — ВОЛНА ОТ НАЖАТИЯ (MD3, адаптировано под мышь)
//
// ДРАЙВЕР — свойство pressed, привязанное к состоянию
// нажимаемого элемента (control.pressed у контролов,
// mouseArea.pressed у иконок). Никаких обработчиков ввода:
// PointHandler в Qt 6.11 не активируется мышью (проверено
// на платформах offscreen и VNC), а перехватчики событий
// мешают доставке клика. Это паттерн Material-стиля самого
// Qt (Ripple { pressed: control.pressed }).
//
// Бонус: нажатие с клавиатуры (Space/Enter) тоже даёт волну.
//
// ФОРМА ВОЛНЫ (растёт из центра до ПОЛНОГО размера элемента,
// никогда не вылезает за его пределы — обрезка не нужна):
//   rippleShape: 0 (круг) — диаметр waveDiameter, или min(width,
//     height): квадратные зоны (иконки, чекбоксы) покрываются
//     целиком;
//   rippleShape: 1 (по форме элемента) — волна повторяет форму
//     самого элемента и вырастает ровно до его размеров: у кнопок-
//     стадионов это пилюля во всю ширину кнопки (radius = height/2
//     или rippleRadius), у трека переключателя — пилюля трека.
//
// Тайминги — десктопные: рост 220 мс, вспышка 80 мс,
// затухание 160 мс; всё оканчивается меньше чем за 400 мс.
// ============================================================

Item {
    id: host

    // Цвет волны: на залитых кнопках — цвет текста кнопки,
    // на прозрачных — бренд. Альфа уже учтена в цвете.
    property color rippleColor: AppTheme.isDark ? Qt.rgba(1, 1, 1, 0.14)
                                                : Qt.rgba(0, 0, 0, 0.12)

    // Можно запретить волны точечно (например, для ghost-кнопок)
    property bool rippleEnabled: true

    // ДРАЙВЕР: true, пока элемент нажат (кнопка удерживается).
    // Привязывается к control.pressed / mouseArea.pressed хоста.
    property bool pressed: false

    property int rippleShape: 0        // 0 — круг, 1 — по форме элемента
    property real waveDiameter: -1     // диаметр круга; -1 = min(width, height)
    property real rippleRadius: -1     // радиус углов волны формы; -1 = height/2

    readonly property real waveD: Math.max(12,
        waveDiameter > 0 ? waveDiameter : Math.min(width, height))

    // стартовый масштаб: круг — из точки; форма — с высоты ~14px,
    // чтобы пилюля не начиналась «спичкой» на широких кнопках
    readonly property real startScale: rippleShape === 1
        ? Math.max(0.12, 14 / Math.max(height, 1))
        : 0.06

    // Ручной запуск (тесты, нестандартные хосты)
    function startRipple() {
        rippleAnim.restart()
    }

    onPressedChanged: {
        if (pressed && rippleEnabled && width > 0)
            startRipple()
    }

    anchors.fill: parent
    clip: true

    Rectangle {
        id: wave
        objectName: "rippleWave"
        anchors.centerIn: parent
        // круг: min-сторона (покрывает квадратные зоны целиком);
        // форма: точные размеры элемента, радиус — как у элемента
        width: host.rippleShape === 1 ? host.width : host.waveD
        height: host.rippleShape === 1 ? host.height : host.waveD
        radius: host.rippleShape === 1
                ? (host.rippleRadius >= 0 ? host.rippleRadius : host.height / 2)
                : host.waveD / 2
        color: host.rippleColor
        opacity: 0
        scale: host.startScale
    }

    SequentialAnimation {
        id: rippleAnim
        objectName: "rippleAnim"
        ParallelAnimation {
            NumberAnimation { target: wave; property: "scale"; from: host.startScale; to: 1; duration: 220; easing.type: Easing.OutQuad }
            NumberAnimation { target: wave; property: "opacity"; from: 0; to: 1; duration: 80; easing.type: Easing.OutQuad }
        }
        NumberAnimation { target: wave; property: "opacity"; from: 1; to: 0; duration: 160; easing.type: Easing.InQuad }
    }
}
