import QtQuick

// ═══════════════════════════════════════════════════════════════════
// ВДАВЛЕННАЯ РАМКА — точечный неоморфизм (soft UI)
//
// Компаньон AppSoftShadow: «втапливает» элемент в поверхность —
// сверху-слева тёмная кромка, снизу-справа светлая. Чекбоксы и
// прочие «ямки». Тоже готовый PNG (inset_l1), стоит почти ноль.
//
// Использование (объявлять ПЕРВЫМ дочерним элементом Rectangle,
// чтобы контент рисовался поверх рамки):
//     Rectangle {
//         color: AppTheme.bgSurface
//         AppInsetShadow { level: 1 }
//         ...
//     }
// ═══════════════════════════════════════════════════════════════════
BorderImage {
    id: root

    // Пока один уровень: чекбоксы и небольшие «ямки»
    property int level: 1

    // Рамка чуть выходит за края элемента (запас, вшитый в PNG)
    readonly property var _pads:   [0, 2]
    // Неломаемая рамка картинки (углы не растягиваются)
    readonly property var _insets: [0, 8]

    anchors.fill: parent
    anchors.leftMargin:   -_pads[level]
    anchors.rightMargin:  -_pads[level]
    anchors.topMargin:    -_pads[level]
    anchors.bottomMargin: -_pads[level]

    source: AppTheme.isDark
            ? "../shadows/inset_l" + level + "_dark.png"
            : "../shadows/inset_l" + level + ".png"
    border.left:   _insets[level]
    border.right:  _insets[level]
    border.top:    _insets[level]
    border.bottom: _insets[level]
    smooth: true
    cache: true
}
