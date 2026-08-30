import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

AppDialog {
    id: root
    width: 380

    title: "Денежная выплата"
    acceptText: "Сохранить"
    acceptVariant: "primary"

    // Массив добавленных компенсаций: [{ unit: "hours", label: "Ночные", amount: 8 }]
    property var activeComps: []
    property int editCompId: 0

    // Функция сброса окна при новом открытии
    function openNew(callerItem, mouseX, mouseY) {
        root.editCompId = 0
        root.activeComps = []
        moneyOrderInput.text = ""
        moneyCommentInput.text = ""
        moneyErrorMsg.visible = false
        moneyDateInput.selectedDate = new Date().toISOString().split('T')[0]
        root.showAt(callerItem, mouseX, mouseY)
    }

    function openEdit(compData, callerItem, mouseX, mouseY) {
        root.editCompId = compData.id
        root.activeComps = [{ unit: compData.unit, label: compData.type, amount: compData.raw_amount }]
        moneyOrderInput.text = compData.order_no
        moneyCommentInput.text = compData.comment
        moneyErrorMsg.visible = false
        moneyDateInput.selectedDate = compData.date.split(".").reverse().join("-")
        root.showAt(callerItem, mouseX, mouseY)
    }

    // ==========================================
    // 1. БЛОК ПИЛЮЛЬ (ЧТО ПЛАТИМ) - ТЕПЕРЬ СВЕРХУ
    // ==========================================
    Column {
        width: parent.width
        spacing: AppTheme.spaceS

        Text { 
            text: "Что оплачиваем:"
            color: AppTheme.textSecondary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightBold 
        }

        Flow {
            width: parent.width
            spacing: AppTheme.spaceS

            // Вывод добавленных пилюль
            Repeater {
                model: root.activeComps
                AppPill {
                    removable: true
                    text: modelData.label + ": " + modelData.amount
                    
                    // По клику на саму пилюлю (не на крестик) - открываем редактирование
                    onClicked: {
                        compEditDialog.isEditing = true
                        compEditDialog.targetUnit = modelData.unit
                        compEditDialog.targetLabel = modelData.label
                        compEditDialog.amountText = modelData.amount.toString()
                        compEditDialog.showCentered()
                    }
                    
                    // Удаление
                    onRemoveClicked: {
                        let arr = root.activeComps.slice()
                        arr.splice(index, 1)
                        root.activeComps = arr
                    }
                }
            }

            // Кнопка "Добавить" видна, только если добавлены не все 3 типа
            AppPill {
                id: addCompBtn
                isAction: true
                text: "Добавить"
                visible: root.activeComps.length < 3 && root.editCompId === 0
                onClicked: {
                    compEditDialog.isEditing = false
                    compEditDialog.amountText = ""
                    compEditDialog.showAt(addCompBtn, addCompBtn.width / 2, addCompBtn.height)
                }
            }
        }
    }

    // Разделитель
    Rectangle { 
        width: parent.width
        height: 1
        color: AppTheme.borderDivider 
    }

    // ==========================================
    // 2. ОСНОВНЫЕ ПОЛЯ (По одному в строке)
    // ==========================================
    AppDateField { 
        id: moneyDateInput
        width: parent.width
        label: "Дата приказа"
        isRequired: true
    }

    AppTextField { 
        id: moneyOrderInput
        width: parent.width
        label: "Номер приказа"
        isRequired: true
        placeholderText: "123 л/с" 
    }

    AppTextField { 
        id: moneyCommentInput
        width: parent.width
        label: "Комментарий"
        placeholderText: "Например: Оплата за Рождество" 
    }

    // Ошибка в окне (а не тост): обязательные поля, пустой список выплат
    Text {
        id: moneyErrorMsg
        visible: false
        width: parent.width
        color: AppTheme.accentDanger
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeSmall
        wrapMode: Text.WordWrap
    }

    // ==========================================
    // 3. СОХРАНЕНИЕ
    // ==========================================
    onAccepted: {
        if (moneyOrderInput.text.trim() === "" || !moneyDateInput.selectedDate) {
            moneyErrorMsg.text = "Ошибка: Укажите дату и номер приказа"
            moneyErrorMsg.visible = true
            root.shake()
            return
        }

        if (root.activeComps.length === 0) {
            moneyErrorMsg.text = "Ошибка: Добавьте хотя бы одну выплату"
            moneyErrorMsg.visible = true
            root.shake()
            return
        }
        moneyErrorMsg.visible = false
        // ------------------------------

        if (root.editCompId > 0) {
            let comp = root.activeComps[0]
            backend.updateMoneyComp(root.editCompId, comp.unit, comp.amount.toString(), moneyOrderInput.text, moneyCommentInput.text)
        } else {
            let compsJson = JSON.stringify(root.activeComps)
            backend.saveMoneyCompList(compsJson, moneyOrderInput.text, moneyDateInput.selectedDate, moneyCommentInput.text)
        }
        root.close()
    }

    // ==========================================
    // 4. САБ-ОКНО ДОБАВЛЕНИЯ / РЕДАКТИРОВАНИЯ
    // ==========================================
    AppDialog {
        id: compEditDialog
        parent: Overlay.overlay 
        width: 320
        title: isEditing ? "Редактировать" : "Что списываем?"
        acceptText: "Применить"

        property bool isEditing: false
        property string targetUnit: ""
        property string targetLabel: ""
        property alias amountText: amtInput.text

        // Вычисляем доступные типы (исключаем те, что уже добавлены)
        property var availableTypes: {
            let all = [
                { text: "Ночные (ч)", value: "hours" },
                { text: "Сверх нормы (ч)", value: "overtime" },
                { text: "Дни", value: "days" }
            ]
            if (isEditing) return []
            let used = root.activeComps.map(c => c.unit)
            return all.filter(t => !used.includes(t.value))
        }

        onOpened: {
            errorMsg.visible = false
            prevYearCheck.checked = false // Сбрасываем при открытии
            if (!isEditing && availableTypes.length > 0) {
                typeCombo.model = availableTypes
                typeCombo.currentIndex = 0
            }
        }

        Column {
            width: parent.width
            spacing: AppTheme.spaceM

            // Комбобокс виден только при создании
            AppComboBox {
                id: typeCombo
                width: parent.width
                visible: !compEditDialog.isEditing
                label: "Вид компенсации"
                textRole: "text"
                valueRole: "value"
            }

            // Чекбокс "Предыдущий год"
            AppCheckBox {
                id: prevYearCheck
                text: "Списать из предыдущего года"
                visible: !compEditDialog.isEditing
                width: parent.width
            }

            // При редактировании показываем просто текст
            Text {
                visible: compEditDialog.isEditing
                text: "Вид: " + compEditDialog.targetLabel
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBodyLarge
                font.weight: AppTheme.weightBold
            }

            AppTextField {
                id: amtInput
                width: parent.width
                label: "Количество:"
                validator: RegularExpressionValidator { regularExpression: /^[1-9][0-9]*$/ }
                onTextChanged: errorMsg.visible = false
            }

            Text {
                id: errorMsg
                width: parent.width
                visible: false
                color: AppTheme.accentDanger
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                wrapMode: Text.WordWrap
            }
        }

        onAccepted: {
            let val = parseInt(amtInput.text)
            if (isNaN(val) || val <= 0) {
                errorMsg.text = "Ошибка: Укажите количество"
                errorMsg.visible = true
                compEditDialog.shake()
                return
            }

            // МАГИЯ: Берем год из даты приказа, но если стоит галка - отнимаем 1
            let baseYear = parseInt(moneyDateInput.selectedDate.split('-')[0]) || new Date().getFullYear()
            let targetYear = prevYearCheck.checked ? baseYear - 1 : baseYear
            
            // Спрашиваем бэкенд остатки именно за выбранный год
            let balances = backend.getAvailableBalances(targetYear)
            
            let currentUnit = isEditing ? targetUnit : typeCombo.currentValue
            let currentLabel = isEditing ? targetLabel : typeCombo.currentText
            
            let maxAllowed = balances[currentUnit] || 0

            if (val > maxAllowed) {
                compEditDialog.shake()
                errorMsg.text = "Ошибка: Доступно максимум " + maxAllowed + " (в " + targetYear + " г.)"
                errorMsg.visible = true
                return
            }

            let arr = root.activeComps.slice()
            if (isEditing) {
                for (let i = 0; i < arr.length; i++) {
                    if (arr[i].unit === targetUnit) {
                        arr[i].amount = val
                        break
                    }
                }
            } else {
                arr.push({ unit: currentUnit, label: currentLabel, amount: val, usePrevYear: prevYearCheck.checked })
            }
            
            root.activeComps = arr
            compEditDialog.close()
        }
    }
}