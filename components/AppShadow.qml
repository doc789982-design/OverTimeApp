import QtQuick

// ═══════════════════════════════════════════════════════════════════
// ЛЁГКАЯ ТЕНЬ-КАРТИНКА (как в Telegram/Discord)
//
// Раньше тени делались через DropShadow: видеокарта КАЖДЫЙ КАДР
// рисовала элемент на невидимый холст, размывала его (25 проходов
// смешивания на точку!) и клеила обратно. На слабых видеокартах
// это главный источник лагов.
//
// Теперь тень — это заранее нарисованный PNG (папка shadows/),
// который просто растягивается под элементом. Стоимость — как у
// обычной картинки, т.е. почти ноль. Выглядит идентично.
//
// Использование (вместо layer.enabled + layer.effect: DropShadow):
//     Rectangle {
//         AppShadow { level: 4 }   // уровни 1..5, те же что были
//     }
// ═══════════════════════════════════════════════════════════════════
BorderImage {
    id: root

    // Уровень тени 1..5 — соответствует старым AppTheme.shadowL#Blur
    property int level: 1

    // Сдвиг тени вниз (отрицательный = вверх, как у нижней шторки)
    property int yOffset: _offsets[level]

    // Насколько картинка выступает за края элемента (запас на размытие)
    readonly property var _pads:    [0,  8, 14, 18, 26, 34]
    // Неломаемая рамка картинки (углы не растягиваются)
    readonly property var _insets:  [0, 16, 22, 26, 42, 42]
    // Стандартные вертикальные сдвиги (те же, что были у DropShadow)
    readonly property var _offsets: [0,  1,  3,  4,  8, 12]

    // Та же прозрачность, что была у AppTheme.shadowColor
    opacity: AppTheme.isDark ? 0.40 : 0.12

    z: -1  // рисуемся ПОД родителем
    anchors.fill: parent
    anchors.leftMargin:   -_pads[level]
    anchors.rightMargin:  -_pads[level]
    anchors.topMargin:    -_pads[level] + yOffset
    anchors.bottomMargin: -_pads[level] - yOffset

    source: "../shadows/shadow_l" + level + ".png"
    border.left:   _insets[level]
    border.right:  _insets[level]
    border.top:    _insets[level]
    border.bottom: _insets[level]
    smooth: true
    cache: true
}
