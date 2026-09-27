import QtQuick

// ============================================================
// ИНЛАЙН-ТУЛТИП — карточка обычным Item'ом, а не попапом.
//
// Зачем: внутри попапов (меню дня, комбобоксы) ToolTip-попап
// рендерится в overlay главного окна ПОД попапом-хостом —
// пользователь видит только краешек, вылезающий за панель
// (проверено на offscreen: зона тултипа перекрывается меню
// независимо от z). Инлайн-карточка с высоким z рисуется
// поверх содержимого хоста и видна целиком в пределах окна
// приложения. Для тултипов в меню — ровно то, что нужно.
//
// Стиль и поведение — как у AppToolTip: карточка bgElevated
// с рамкой и тенью, задержка 350 мс, мгновенное скрытие.
// ============================================================

Item {
    id: root

    property string text: ""
    property bool isVisible: false
    property int delayMs: 350   // наведение должно быть осознанным

    width: tipCard.width
    height: tipCard.height
    z: AppTheme.zTooltip
    opacity: 0
    visible: opacity > 0.01

    onIsVisibleChanged: {
        if (root.isVisible) showDelayTimer.restart()
        else { showDelayTimer.stop(); root.opacity = 0 }
    }

    Timer {
        id: showDelayTimer
        interval: Math.max(0, root.delayMs)
        onTriggered: root.opacity = 1
    }

    Rectangle {
        id: tipCard
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

    Behavior on opacity { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }
}
