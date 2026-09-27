import QtQuick
import QtQuick.Controls

// ============================================================
// СКРОЛЛБАР В СТИЛЕ ПРОГРАММЫ (раньше был дефолтный серый)
// По MD3: без видимой дорожки, тонкий округлый бегунок,
// подсветка при наведении/нажатии.
// ============================================================
ScrollBar {
    id: root

    implicitWidth: 8
    implicitHeight: 8

    background: null

    contentItem: Rectangle {
        radius: width / 2
        implicitWidth: 8
        implicitHeight: 24
        color: root.pressed ? AppTheme.textSecondary
               : root.hovered ? AppTheme.textTertiary
               : AppTheme.borderDivider
        opacity: root.pressed ? 1.0 : 0.85

        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro } }
    }
}
