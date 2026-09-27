import QtQuick
import QtQuick.Controls

// ============================================================
// СКРОЛЛБАР В СТИЛЕ ПРОГРАММЫ
// MD3 (desktop): бегунок без дорожки, скруглённый, заметный;
// при наведении индикатор расширяется и темнеет.
// Сайт использования обязан зарезервировать желоб
// (AppTheme.scrollGutter), чтобы бегунок лежал РЯДОМ с контентом,
// а не поверх него.
// ============================================================
ScrollBar {
    id: root

    implicitWidth: AppTheme.scrollGutter
    implicitHeight: AppTheme.scrollGutter

    background: null

    contentItem: Rectangle {
        radius: width / 2
        implicitWidth: 6
        implicitHeight: 24

        // центрируем в желобе (вертикальный скроллбар)
        x: (parent ? (parent.width - width) / 2 : 0)

        // MD3: при наведении индикатор расширяется
        width: (root.hovered || root.pressed) ? 10 : 6
        Behavior on width { NumberAnimation { duration: AppTheme.durFast; easing.type: Easing.OutCubic } }

        color: root.pressed ? AppTheme.accentBrand
               : root.hovered ? AppTheme.textSecondary
               : AppTheme.textTertiary
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }
}
