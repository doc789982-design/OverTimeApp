import QtQuick

// ============================================================
// ПЕРЕКЛЮЧАТЕЛЬ-ПИЛЮЛЯ (вместо AppSwitch) — по CSS-образцу.
//
// Хореография образца (главное, что видно глазу):
//   КРУЖОК РАСШИРЯЕТСЯ. Ножка старого состояния сама становится
//   волной: круг РАСТЁТ от своего места (600 мс) и заливает
//   пилюлю цветом нового состояния. Прежнее поле в этот же
//   миг мгновенно сжимается в кружок — но оно ПОД волной и
//   проявляется как новая ножка только В КОНЦЕ анимации,
//   когда z-порядок меняется (в образце — transition
//   z-index 0s .6s: смена слоя с задержкой 600 мс).
//
// Фон пилюли держит прошлый цвет до 80% анимации (480 мс),
// чтобы у растущего круга не было видно швов по краям
// (keyframes changeColor 80%/80.01% образца).
//
// Пропорции образца 252×126 (2:1): ножка 80/126 высоты,
// отступ 23/126. По умолчанию пилюля 64×32 — те же пропорции,
// «округлая», как в образце. Анимации включаются только после
// постройки компонента: при открытии окна свитч встаёт в своё
// состояние мгновенно, без холостого «разгона» волны.
//
// Совместим со старым AppSwitch: text (подпись справа),
// checked, сигнал toggled() (только от клика человека).
// Цвета фиксированы палитрой темы: чернильный #2D3B45 и
// #FFFFFF — переключатель сам показывает «тёмное/светлое».
// ============================================================
Item {
    id: root

    property bool checked: false
    signal toggled()

    property string text: ""

    // Размер пилюли (трека); подпись добавляет ширину сама
    property real pillWidth: 64
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
    // Во сколько раз волна закрывает пилюлю: с перехлёстом,
    // как в образце (80 × 4.8 = 384 при пилюле 252)
    readonly property real coverScale: Math.max(pillWidth, pillHeight) / knobD * 1.5

    // Волна в полёте: растущий круг СВЕРХУ, ножка ПОД ним и
    // проявляется в конце (см. шапку)
    property bool waveRunning: false

    // «Водители» масштабов волн: растущий круг анимируется
    // императивно (NumberAnimation), сжимающийся прыгает
    // мгновенно. Behavior тут НЕ подходит: длительность
    // перевязывается в тот же тик, что и цель — поведение
    // зависит от порядка вычисления связей (гонка)
    property real darkScale: 1
    property real lightScale: 1

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
            // покой: выкл — поле (под ножкой), вкл — ножка (над волной);
            // в полёте: растущая волна всегда сверху
            z: root.checked ? (root.waveRunning ? 1 : 2)
                            : (root.waveRunning ? 2 : 1)
            // сжавшийся круг уходит ПОД волну (мгновенность —
            // в драйвере, см. darkScale/lightScale)
            scale: root.darkScale
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
            z: root.checked ? (root.waveRunning ? 2 : 1)
                            : (root.waveRunning ? 1 : 2)
            scale: root.lightScale
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
    // сразу светлый. Внутренняя реакция — через Connections на
    // ребёнке: прямой onCheckedChanged экземпляра её бы затёр.
    Timer {
        id: bgHold
        interval: 480
        onTriggered: pill.color = root.lightColor
    }
    // Смена слоёв — в КОНЦЕ полёта волны (transition z-index 0s .6s)
    Timer {
        id: zFlip
        interval: 600
        onTriggered: root.waveRunning = false
    }
    // Рост волны — 600 мс (transition .6s ease образца)
    NumberAnimation {
        id: growDark
        target: root
        property: "darkScale"
        from: 1
        to: root.coverScale
        duration: 600
        easing.type: Easing.InOutQuad
    }
    NumberAnimation {
        id: growLight
        target: root
        property: "lightScale"
        from: 1
        to: root.coverScale
        duration: 600
        easing.type: Easing.InOutQuad
    }
    // Начальное состояние — мгновенно, без анимации (checked мог
    // прийти с бэкенда уже true)
    Component.onCompleted: {
        if (root.checked) {
            root.darkScale = 1
            root.lightScale = root.coverScale
        } else {
            root.darkScale = root.coverScale
            root.lightScale = 1
        }
    }
    Connections {
        target: root
        function onCheckedChanged() {
            root.waveRunning = true
            zFlip.restart()
            if (root.checked) {
                // тёмная мгновенно в ножку (и под волну),
                // светлая растёт и заливает пилюлю
                root.darkScale = 1
                growLight.restart()
                pill.color = root.darkColor
                bgHold.restart()
            } else {
                root.lightScale = 1
                growDark.restart()
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
