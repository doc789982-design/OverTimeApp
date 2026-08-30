import QtQuick
import QtQuick.Controls

AppDialog {
    id: root
    width: 320
    
    title: "Новое подразделение"
    acceptText: "Создать"
    acceptVariant: "primary" 

    AppTextField { 
        id: startupNewDbNameInput
        width: parent.width
        label: "Название отдела"
        isRequired: true
        placeholderText: "1 Отдел"
    }

    // Ошибка в окне (а не тост)
    Text {
        id: createDbErrorMsg
        visible: false
        width: parent.width
        color: AppTheme.accentDanger
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeSmall
        wrapMode: Text.WordWrap
    }

    onAboutToShow: createDbErrorMsg.visible = false

    onAccepted: {
        if (startupNewDbNameInput.text.trim() === "") {
            createDbErrorMsg.text = "Ошибка: Введите название отдела"
            createDbErrorMsg.visible = true
            root.shake()
            return
        }
        createDbErrorMsg.visible = false
        backend.createNewDatabase(startupNewDbNameInput.text)
        startupNewDbNameInput.text = ""
        root.close()
    }
}