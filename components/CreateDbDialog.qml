import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО СОЗДАНИЯ НОВОЙ БАЗЫ: кроме названия отдела сразу спрашивает
// ответственное лицо — те же поля, что потом живут в настройках
// («Данные подразделения»: должность, звание, ФИО). Данные идут в
// шапки и подписи экспортируемых документов — теперь их не нужно
// до заполнять после создания. Поля не обязательны: базу можно
// создать и без них, добрать позже в настройках.
// ============================================================
AppDialog {
    id: root
    width: 500
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

    // ── Ответственное лицо ──
    Text {
        width: parent.width
        topPadding: AppTheme.spaceS
        text: "Ответственное лицо"
        color: AppTheme.textPrimary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightBold
    }
    RowLayout {
        width: parent.width
        spacing: AppTheme.spaceS
        AppTextField {
            id: dbRespPosition
            Layout.preferredWidth: parent.width * 0.62
            Layout.fillWidth: true
            label: "Должность"
            placeholderText: "Начальник отдела"
        }
        AppTextField {
            id: dbRespRank
            Layout.preferredWidth: parent.width * 0.38
            Layout.fillWidth: true
            label: "Звание"
            placeholderText: "Подполковник полиции"
        }
    }
    RowLayout {
        width: parent.width
        spacing: AppTheme.spaceS
        AppTextField {
            id: dbRespLast
            Layout.preferredWidth: parent.width / 3
            Layout.fillWidth: true
            label: "Фамилия"
        }
        AppTextField {
            id: dbRespFirst
            Layout.preferredWidth: parent.width / 3
            Layout.fillWidth: true
            label: "Имя"
        }
        AppTextField {
            id: dbRespMid
            Layout.preferredWidth: parent.width / 3
            Layout.fillWidth: true
            label: "Отчество"
        }
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
            startupNewDbNameInput.flashError()
            root.shake()
            return
        }
        createDbErrorMsg.visible = false
        backend.createNewDatabase(
            startupNewDbNameInput.text,
            dbRespPosition.text, dbRespRank.text,
            dbRespLast.text, dbRespFirst.text, dbRespMid.text)
        startupNewDbNameInput.text = ""
        dbRespPosition.text = ""
        dbRespRank.text = ""
        dbRespLast.text = ""
        dbRespFirst.text = ""
        dbRespMid.text = ""
        root.close()
    }
}
