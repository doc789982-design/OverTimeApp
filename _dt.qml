
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow { width: 1200; height: 800; visible: true
    AppUI.MoneyOrderDialog { id: dlg }
    Component.onCompleted: dlg.openNew() }
