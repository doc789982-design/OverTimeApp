import QtQuick
import QtQuick.Controls

// ============================================================
// ПРОСТОЕ ТЕКСТОВОЕ ПОЛЕ ФОРМЫ: базовая механика — в AppFieldBase,
// здесь только защита числовых полей: буквы физически невозможно
// ввести.
// ============================================================
AppFieldBase {
    id: root

    property bool numericOnly: false   // true = поле принимает только цифры

    validator: root.numericOnly ? digitsOnly : null
    RegularExpressionValidator {
        id: digitsOnly
        regularExpression: /^[0-9]{0,5}$/
    }
}
