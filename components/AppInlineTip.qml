import QtQuick
import QtQuick.Controls

// ============================================================
// ИНЛАЙН-ТУЛТИП — карточка обычным Item'ом, а не попапом.
//
// Зачем: ToolTip-попап, объявленный внутри другого попапа
// (меню дня), рендерится в overlay главного окна ПОД
// попапом-хостом, а содержимое самого меню клипуется его
// ListView (clip: true в стиле Basic). Поэтому карточка
// живёт там, где её ничего не режет: в overlay главного
// окна (parent: ApplicationWindow.overlay), с z выше меню.
//
// ПОЗИЦИОНИРОВАНИЕ: showAt(item, text) ставит карточку над
// элементом item в координатах родителя (overlay); ховер
// кнопки меню вызывает showAt/hideTip. АНИМАЦИЯ — один в
// один как у AppToolTip: выезд 4px + фейд 150 мс, уход —
// быстрый фейд. Задержка появления — 350 мс.
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

    // Показать над элементом item (в координатах родителя этой карточки)
    function showAt(item, txt) {
        root.text = txt
        var p = item.mapToItem(root.parent, item.width / 2, 0)
        root.x = p.x - root.width / 2
        root.y = p.y - root.height - AppTheme.spaceXS
        root.isVisible = true
    }

    function hideTip() {
        root.isVisible = false
    }

    // Перенос карточки в overlay главного окна (поверх всех попапов;
    // содержимое меню клипуется его ListView). Вызывается хостом,
    // когда попап-меню уже открыт и окно доступно.
    function attachToOverlay() {
        var ov = Overlay.overlay
        if (ov && root.parent !== ov)
            root.parent = ov
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
