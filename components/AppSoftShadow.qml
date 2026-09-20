import QtQuick

// ═══════════════════════════════════════════════════════════════════
// ВЫПУКЛАЯ ТЕНЬ-КАРТИНКА — точечный неоморфизм (soft UI)
//
// Младший брат AppShadow для мест, где элемент должен быть «выдавлен»
// из фона: карточка выбранного сотрудника, плашка иконки в балансовых
// карточках, крупные панели. В одном PNG сразу две тени — светлая
// сверху-слева и тёмная снизу-справа; для тёмной темы свой файл.
//
// Классические AppShadow (диалоги, меню, всплывающие окна) не меняются.
//
// Использование:
//     Rectangle {
//         AppSoftShadow { level: 2 }   // 1 — мелочь, 2 — крупные панели
//     }
// ═══════════════════════════════════════════════════════════════════
BorderImage {
    id: root

    // Уровень: 1 — карточки и плашки, 2 — крупные панели
    property int level: 1

    // Насколько картинка выступает за края элемента (запас на размытие)
    readonly property var _pads:   [0, 12, 18]
    // Неломаемая рамка картинки (углы не растягиваются)
    readonly property var _insets: [0, 24, 30]

    opacity: 1.0
    z: -1  // рисуемся ПОД родителем
    anchors.fill: parent
    anchors.leftMargin:   -_pads[level]
    anchors.rightMargin:  -_pads[level]
    anchors.topMargin:    -_pads[level]
    anchors.bottomMargin: -_pads[level]

    source: AppTheme.isDark
            ? "../shadows/soft_l" + level + "_dark.png"
            : "../shadows/soft_l" + level + ".png"
    border.left:   _insets[level]
    border.right:  _insets[level]
    border.top:    _insets[level]
    border.bottom: _insets[level]
    smooth: true
    cache: true
}
