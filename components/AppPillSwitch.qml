import QtQuick

// ============================================================
// БОЛЬШОЙ ПЕРЕКЛЮЧАТЕЛЬ-«ПИЛЮЛЯ» (по мотивам CSS-свотча:
// две волны-круга расходятся и закрывают пилюлю, вторая
// мгновенно сжимается в кружок-ногу).
//
// Механика оригинала (252×126, круги 80px, 600 мс):
//   - «Волна» — круг в ножке; включаясь, РАСТЁТ (scale 1 → 4.8)
//     за 600 мс и закрывает всю пилюлю своим цветом;
//   - вторая волна в этот момент МГНОВЕННО сжимается в кружок —
//     это новая «нога» переключателя;
//   - фон пилюли держит прошлый цвет до 80% анимации (480 мс),
//     чтобы у растущего круга не было видно швов по краям;
//   - z-порядок меняется в момент клика: маленький круг (нога)
//     всегда сверху.
//
// Цвета — палитра нашей темы (не зависимости от тёмной/светлой):
// тёмная сторона — чернильный #2D3B45 (textPrimary светлой
// темы), светлая — #FFFFFF (bgElevated). Оба можно переопределить
// свойствами darkColor / lightColor.
//
// Внешний вид: пилюля с большой мягкой тенью (AppShadow level 4;
// в тёмной теме тень скрыта по правилам дизайн-системы).
// ============================================================
Item {
    id: root

    width: 252
    height: width / 2

    property bool checked: false
    signal toggled()

    // Цвета сторон (палитра темы, фиксированы — переключатель сам
    // показывает «светлое/тёмное» и не должен меняться с темой)
    property color darkColor: "#2D3B45"
    property color lightColor: "#FFFFFF"

    // Геометрия пропорциональна высоте (80/126 и 23/126 — как в
    // оригинале), поэтому компонент можно уменьшать шириной
    readonly property real knobD: height * 80 / 126      // диаметр круга
    readonly property real knobM: height * 23 / 126      // отступ от края
    readonly property real coverScale: 4.8 * width / 252 // во сколько раз
                                                          // волна закрывает пилюлю

    opacity: enabled ? 1.0 : AppTheme.alphaDisabled

    function toggle() {
        root.checked = !root.checked
        root.toggled()
    }

    // Большая мягкая тень (уровень 4); в тёмной теме AppShadow
    // сама становится невидимой
    AppShadow { level: 4 }

    // Пилюля; фон — «светлый», волны его полностью закрывают.
    // clip = overflow: hidden у оригинала.
    Rectangle {
        id: pill
        anchors.fill: parent
        radius: height / 2
        color: root.lightColor
        clip: true

        // Тёмная волна (ножка слева)
        Rectangle {
            id: rippleDark
            x: root.knobM
            y: root.knobM
            width: root.knobD
            height: root.knobD
            radius: width / 2
            color: root.darkColor
            z: root.checked ? 2 : 1
            scale: root.checked ? 1 : root.coverScale
            // растёт 600 мс, сжимается мгновенно (как transition
            // transform 0s у оригинала)
            Behavior on scale {
                NumberAnimation {
                    duration: root.checked ? 0 : 600
                    easing.type: Easing.InOutQuad
                }
            }
        }

        // Светлая волна (ножка справа)
        Rectangle {
            id: rippleLight
            x: parent.width - root.knobM - root.knobD
            y: root.knobM
            width: root.knobD
            height: root.knobD
            radius: width / 2
            color: root.lightColor
            z: root.checked ? 1 : 2
            scale: root.checked ? root.coverScale : 1
            Behavior on scale {
                NumberAnimation {
                    duration: root.checked ? 600 : 0
                    easing.type: Easing.InOutQuad
                }
            }
        }
    }

    // Фон пилюли: при включении мгновенно тёмный и ОСТАЁТСЯ им до
    // 80% анимации (480 мс), затем светлеет; при выключении —
    // сразу светлый (changeColor 80%/80.01% у оригинала)
    Timer {
        id: bgHold
        interval: 480
        onTriggered: pill.color = root.lightColor
    }
    onCheckedChanged: {
        if (checked) {
            pill.color = root.darkColor
            bgHold.restart()
        } else {
            pill.color = root.lightColor
            bgHold.stop()
        }
    }

    HoverHandler { cursorShape: Qt.PointingHandCursor }
    MouseArea {
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        onClicked: root.toggle()
    }
}
