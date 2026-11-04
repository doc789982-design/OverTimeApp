import QtQuick
import QtQuick.Controls

// ============================================================
// ПРОСТОЕ ТЕКСТОВОЕ ПОЛЕ ФОРМЫ: базовая механика — в AppFieldBase,
// здесь только защита числовых полей: буквы физически невозможно
// ввести. Минус разрешён только там, где поле явно попросило
// allowNegative (сейчас — начальные сверхурочные сотрудника).
// ============================================================
AppFieldBase {
    id: root

    property bool numericOnly: false   // true = поле принимает только цифры
    property bool allowNegative: false // для numericOnly: можно ввести минус

    validator: root.numericOnly
               ? (root.allowNegative ? signedDigitsOnly : digitsOnly)
               : null
    RegularExpressionValidator {
        id: digitsOnly
        regularExpression: /^[0-9]{0,5}$/
    }
    RegularExpressionValidator {
        id: signedDigitsOnly
        regularExpression: /^-?[0-9]{0,5}$/
    }
}
