import QtQuick
import QtQuick.Controls

AppDialog {
    id: root
    width: 350

    property int targetEmpId: 0
    property int targetGroupId: 0

    title: "Официальный перевод"
    acceptText: "Перевести"

    Text {
        text: "Норма и переработки пересчитаются с учетом указанной даты."
        color: AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        width: parent.width
        wrapMode: Text.WordWrap
    }

    AppDateField { 
        id: transferDateInput
        width: parent.width
        label: "Дата перевода:"
        selectedDate: new Date().toISOString().split('T')[0]
    }

    onAccepted: {
        backend.officialTransferEmployee(root.targetEmpId, root.targetGroupId, transferDateInput.selectedDate)
        root.close()
    }
}