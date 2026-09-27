import QtQuick

// ============================================================
// ИНЛАЙН-ТУЛТИП — карточка обычным Item'ом, а не попапом.
//
// Зачем: внутри попапов (меню дня) ToolTip-попап рендерится
// в overlay главного окна ПОД попапом-хостом — пользователь
// видит только краешек за панелью. Инлайн-карточка с высоким
// z рисуется поверх содержимого хоста. Хост-попап при этом
// должен быть Item-попапом (popupType: Popup.Item), иначе
// нативное окно меню обрежет карточку своими краями.
//
// АНИМАЦИЯ — один в один как у AppToolTip: карточка всплывает
// фейдом 150 мс и выездом на 4px (снизу вверх для тултипа над
// элементом, сверху вниз для dropDown), уход — быстрый фейд.
// Задержка появления — 350 мс, скрытие мгновенное.
//
// Состояние держит один флаг shown: все свойства declarative,
// никаких анимаций с явными from — те дёргают свойства ещё
// при создании компонента.
// ============================================================

Item {
    id: root

    property string text: ""
    property bool isVisible: false
    property bool dropDown: false   // тултип под элементом — прилетает сверху
    property int delayMs: 350   // наведение должно быть осознанным

    width: tipCard.width
    height: tipCard.height
    z: AppTheme.zTooltip

    property bool shown: false

    opacity: shown ? 1 : 0
    visible: opacity > 0.01
    Behavior on opacity {
        NumberAnimation { duration: root.shown ? AppTheme.durFast : AppTheme.durMicro }
    }

    onIsVisibleChanged: {
        if (root.isVisible) showDelayTimer.restart()
        else { showDelayTimer.stop(); root.shown = false }
    }

    Timer {
        id: showDelayTimer
        interval: Math.max(0, root.delayMs)
        onTriggered: root.shown = true
    }

    Rectangle {
        id: tipCard
        objectName: "inlineTipCard"
        // выезд на 4px: снизу вверх (над элементом), сверху вниз (dropDown)
        y: root.shown ? 0 : (root.dropDown ? -4 : 4)
        Behavior on y { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }

        width: tipText.implicitWidth + AppTheme.spaceM * 2
        height: tipText.implicitHeight + AppTheme.spaceS * 2
        radius: AppTheme.radiusMedium
        color: AppTheme.bgElevated
        border.color: AppTheme.borderDivider
        border.width: 1
        AppShadow { level: 3 }

        Text {
            id: tipText
            anchors.centerIn: parent
            text: root.text
            color: AppTheme.textPrimary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightMedium
        }
    }
}
