import QtQuick
import QtQuick.Controls

AppDialog {
    id: root
    width: 320

    property int targetEmpId: 0
    property string targetReason: "dismissal" 

    title: root.targetReason === "dismissal" ? "Уволить сотрудника" : "Снять со смен"
    acceptText: "Применить"
    acceptVariant: "danger"

    AppDateField { 
        id: endDateInput
        width: parent.width
        label: "Дата события:"
        selectedDate: new Date().toISOString().split('T')[0]
    }

    onAccepted: {
        backend.setEmployeeEndDate(root.targetEmpId, endDateInput.selectedDate, root.targetReason)
        root.close()
    }
}