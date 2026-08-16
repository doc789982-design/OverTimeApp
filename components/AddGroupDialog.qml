import QtQuick
import QtQuick.Controls

AppDialog {
    id: root
    width: 320
    
    title: "Новая группа"
    acceptText: "Создать"

    AppTextField { 
        id: grpNameInput
        width: parent.width
        label: "Название группы"
        isRequired: true // Обязательное поле!
        placeholderText: "Смена 1" // Лаконичная подсказка
    }
    
    AppSwitch {
        id: grpShiftCheck
        text: "Сменный график"
    }
    
    onAccepted: {
        if (grpNameInput.text.trim() === "") {
            root.shake()
            backend.showToast("Ошибка: Введите название группы", "error")
            return
        }
        
        backend.createGroup(grpNameInput.text, grpShiftCheck.checked)
        grpNameInput.text = ""
        root.close()
    }
}