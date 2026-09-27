import QtQuick
import QtQuick.Window
Window { width: 200; height: 100; visible: true
  Text { text: backend.flag ? "OK" : "NULL" }
}