import QtQuick

// ============================================================
// ПЕРЕКЛЮЧАТЕЛЬ-ПИЛЮЛЯ (вместо AppSwitch) — по CSS-образцу:
// волна-круг растёт и закрывает пилюлю, вторая мгновенно
// сжимается в ножку, фон держит прошлый цвет до 80% анимации.
//
// Механика образца (600 мс):
//   - «Волна» — круг в ножке; включаясь, РАСТЁТ (scale 1 →
//     coverScale) за 600 мс и закрывает всю пилюлю;
//   - волна второй стороны в этот же момент МГНОВЕННО
//     сжимается в круг — становится новой ножкой;
//   - фон пилюли держит прошлый цвет 480 мс (80% анимации),
//     чтобы у растущего круга не было видно швов по краям;
//   - ножка всегда над волной.
//
// Совместим со старым AppSwitch: text (подпись справа),
// checked, сигнал toggled() (только от клика человека),
// размеры трека 52×32 по умолчанию — вёрстка не ломается.
// Цвета фиксированы палитрой темы: чернильный #2D3B45 и
// #FFFFFF — переключатель сам показывает «тёмное/светлое».
// ============================================================
Item {
    id: root

    property bool checked: false
    signal toggled()

    property string text: ""

    // Размер пилюли (трека); подпись добавляет ширину сама
    property real pillWidth: 52
    property real pillHeight: 32

    // Цвета сторон (палитра темы; переключатель показывает
    // «светлое/тёмное» и не меняется вместе с темой программы)
    property color darkColor: "#2D3B45"
    property color lightColor: "#FFFFFF"

    // Большая тень — только для крупного варианта (настройки и
    // окна используют маленький размер, тень им не нужна)
    property bool withShadow: false

    // Геометрия пропорциональна пилюле (80/126 и 23/126 —
    // пропорции образца 252×126)
    readonly property real knobD: pillHeight * 80 / 126
    readonly property real knobM: pillHeight * 23 / 126
    // Во сколько раз волна закрывает пилюлю: с запасом, как в
    // образце (80 × 4.8 = 384 при пилюле 252 — рост «с перехлёстом»)
    readonly property real coverScale: Math.max(pillWidth, pillHeight) / knobD * 1.5

    implicitWidth: pillWidth
                   + (text !== "" ? AppTheme.spaceM + label.implicitWidth : 0)
    implicitHeight: 36

    opacity: enabled ? 1.0 : AppTheme.alphaDisabled
    Behavior on opacity { NumberAnimation { duration: AppTheme.durNormal } }

    function toggle() {
        root.checked = !root.checked
        root.toggled()
    }

    AppShadow {
        level: 4
        visible: root.withShadow
    }

    // Пилюля; clip = overflow: hidden образца
    Rectangle {
        id: pill
        width: root.pillWidth
        height: root.pillHeight
        anchors.verticalCenter: parent.verticalCenter
        radius: height / 2
        color: root.lightColor
        clip: true

        // Тёмная волна (ножка слева)
        Rectangle {
            id: rippleDark
            x: root.knobM
            y: (parent.height - root.knobD) / 2
            width: root.knobD
            height: root.knobD
            radius: width / 2
            color: root.darkColor
            z: root.checked ? 2 : 1
            scale: root.checked ? 1 : root.coverScale
            // растёт 600 мс, сжимается мгновенно (transition
            // transform 0s образца)
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
            y: (parent.height - root.knobD) / 2
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

    // Подпись справа от пилюли (как у старого AppSwitch)
    Text {
        id: label
        visible: root.text !== ""
        x: root.pillWidth + AppTheme.spaceM
        anchors.verticalCenter: parent.verticalCenter
        width: implicitWidth
        text: root.text
        color: AppTheme.textPrimary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightMedium
    }

    // Фон пилюли: при включении мгновенно тёмный и ОСТАЁТСЯ им
    // до 80% анимации (480 мс), затем светлеет; при выключении —
    // сразу светлый (keyframes changeColor 80%/80.01% образца).
    // Внутренняя реакция — через Connections на ребёнке: прямой
    // onCheckedChanged экземпляра её бы затёр.
    Timer {
        id: bgHold
        interval: 480
        onTriggered: pill.color = root.lightColor
    }
    Connections {
        target: root
        function onCheckedChanged() {
            if (root.checked) {
                pill.color = root.darkColor
                bgHold.restart()
            } else {
                pill.color = root.lightColor
                bgHold.stop()
            }
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
