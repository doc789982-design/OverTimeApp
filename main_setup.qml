import QtQuick
import QtQuick.Controls
import "components" as AppUI

// Вход мастера установки/удаления: Main.py грузит этот файл с --setup
// и --uninstall (мастер — components/SetupWizard.qml, логика —
// installer.py, заглушка релизного exe — tools/installer_stub.py).
// Вход из корня — тот же приём, что у main.qml: компоненты каталога
// components (модуль с qmldir) видят backend только из такого входа.
ApplicationWindow {
    id: host
    visible: true
    width: 620
    height: 440
    minimumWidth: 560
    minimumHeight: 420
    color: AppUI.AppTheme.bgPanel
    title: backend.mode === "uninstall"
           ? "Удаление OVERTIMETAB"
           : (host.setup.page === "done" && backend.mode !== "uninstall"
              ? "Установка OVERTIMETAB — готово" : "Установка OVERTIMETAB")

    // доступ к состоянию мастера снаружи (заголовок, тесты)
    property alias setup: wizard

    AppUI.SetupWizard {
        id: wizard
        anchors.fill: parent
    }
}
