import QtQuick
import QtQuick.Window
import QtWebEngine

// ============================================================
// ОКНО ВЕБ-ВЕРСИИ ВНУТРИ ПРОГРАММЫ (этап 4 переезда)
//
// Тот же интерфейс, что и в браузере, но живёт как настоящее
// окно программы: своя иконка в панели задач, свой заголовок,
// никакого браузера. Данные считает движок программы —
// сервер поднят в фоновом потоке самого приложения
// (см. Backend._open_web_window в Main.py).
//
// Создаётся ПО ЗАПРОСУ из Python (кнопка «Веб-версия» в справке):
// статический import в main.qml не нужен — программа, где
// WebEngine недоступен, продолжает работать как обычно.
// ============================================================
Window {
    id: root

    property string webUrl: ""

    title: "OVERTIMETAB — веб-версия"
    width: 1320
    height: 900
    minimumWidth: 940
    minimumHeight: 620
    visible: true

    WebEngineView {
        id: view
        anchors.fill: parent
        url: root.webUrl !== "" ? root.webUrl : "about:blank"
    }
}
