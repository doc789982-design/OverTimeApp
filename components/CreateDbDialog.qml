import QtQuick
import QtQuick.Controls

AppDialog {
    id: root
    width: 320
    
    title: "Новое подразделение"
    acceptText: "Создать"
    acceptVariant: "success" 

    AppTextField { 
        id: startupNewDbNameInput
        width: parent.width
        label: "Название отдела"
        isRequired: true
        placeholderText: "1 Отдел"
    }

    onAccepted: {
        if (startupNewDbNameInput.text.trim() === "") {
            root.shake()
            backend.showToast("Введите название!", "error")
            return
        }
        backend.createNewDatabase(startupNewDbNameInput.text)
        startupNewDbNameInput.text = ""
        root.close()
    }
}