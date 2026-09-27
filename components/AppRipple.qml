import QtQuick
import Qt5Compat.GraphicalEffects

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
// Бонус: нажатие с клавиатуры (Space/Enter) тоже даёт волну.
//
// ФОРМА: всегда КРУГ, растущий из центра. Диаметр круга —
// max(width, height), чтобы волна от центра дошла до любого
// края элемента. Обрезка — ПО ФОРМЕ элемента:
//   • аппаратный рендер (Windows) — слой с OpacityMask по
//     скруглению maskRadius: круг никогда не вылезает за
//     скруглённые углы кнопки-стадиона;
//   • software-фолбэк — прямоугольный clip (слои в software
//     мертвы), волна остается внутри габаритов элемента.
//     Слой включается только пока волна живёт — в покое
//     никакого FBO не держим.
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

    // Радиус скругления формы элемента (для маски обрезки):
    // у кнопок-стадионов это height/2, у круглых зон — width/2.
    property real maskRadius: height / 2

    // Диаметр круга: от центра до самого дальнего края элемента
    readonly property real waveD: Math.max(width, height)

    // Аппаратный ли рендер: в software слои (OpacityMask) мертвы,
    // там работает только прямоугольный clip
    readonly property bool hardwareRender: GraphicsInfo.api !== GraphicsInfo.Software

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

    // Маска по форме элемента — только на аппаратном рендере
    // и только пока волна анимируется (в покое слой выключен)
    layer.enabled: hardwareRender && rippleAnim.running
    layer.effect: OpacityMask {
        maskSource: Rectangle {
            width: host.width
            height: host.height
            radius: host.maskRadius
        }
    }

    Rectangle {
        id: wave
        objectName: "rippleWave"
        anchors.centerIn: parent
        width: host.waveD
        height: host.waveD
        radius: host.waveD / 2
        color: host.rippleColor
        opacity: 0
        scale: 0.06
    }

    SequentialAnimation {
        id: rippleAnim
        objectName: "rippleAnim"
        ParallelAnimation {
            NumberAnimation { target: wave; property: "scale"; from: 0.06; to: 1; duration: 220; easing.type: Easing.OutQuad }
            NumberAnimation { target: wave; property: "opacity"; from: 0; to: 1; duration: 80; easing.type: Easing.OutQuad }
        }
        NumberAnimation { target: wave; property: "opacity"; from: 1; to: 0; duration: 160; easing.type: Easing.InQuad }
    }
}
