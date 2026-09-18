import QtQuick

// ═══════════════════════════════════════════════════════════════════
// ВДАВЛЕННАЯ РАМКА — вторая половина неоморфизма (soft UI)
//
// Компаньон AppShadow: если AppShadow «выдавливает» элемент из фона,
// то эта рамка «втапливает» его внутрь. Поле ввода, нажатая кнопка,
// активная ячейка — сверху-слева тёмная кромка, снизу-справа светлая.
//
// Тоже готовый PNG (папка shadows/, файлы inset_l*.png), поэтому
// стоит почти ноль — как обычная картинка.
//
// Использование (объявлять ПЕРВЫМ дочерним элементом Rectangle,
// чтобы контент рисовался поверх рамки):
//     Rectangle {
//         color: AppTheme.bgSurface
//         AppInsetShadow { level: 1 }
//         Text { ... }
//     }
// ═══════════════════════════════════════════════════════════════════
BorderImage {
    id: root

    // Уровень: 1 — поля ввода и мелочи, 2 — крупные вдавленные зоны
    property int level: 1

    // Рамка чуть выходит за края элемента (запас, вшитый в PNG)
    readonly property var _pads:   [0, 2, 2]
    // Неломаемая рамка картинки (углы не растягиваются)
    readonly property var _insets: [0, 10, 16]

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
