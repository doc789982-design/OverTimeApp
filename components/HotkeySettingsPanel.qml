import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

Item {
    id: hotkeyPageRoot
    Layout.fillWidth: true
    Layout.fillHeight: true
    
    property string capturedKey: ""
    property int currentActionTab: 0 // 0: Дежурство, 1: Компенсация, 2: Статус
    property var activeBreaks: []

    ColumnLayout {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceL
        spacing: AppTheme.spaceL
        
        // =========================================================
        // 1. ЗАГОЛОВОК
        // =========================================================
        ColumnLayout {
            spacing: AppTheme.spaceXS
            Layout.fillWidth: true
            Text { text: "Горячие клавиши"; color: AppTheme.textPrimary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeH2; font.weight: AppTheme.weightBold }
            Text { text: "Назначьте действия на клавиатуру для быстрого заполнения табеля."; color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; Layout.fillWidth: true; wrapMode: Text.WordWrap }
        }

        // =========================================================
        // 2. БЛОК ДОБАВЛЕНИЯ НОВОЙ КЛАВИШИ 
        // =========================================================
        Rectangle {
            Layout.fillWidth: true
            implicitHeight: newHkLayout.implicitHeight + (AppTheme.spaceL * 2)
            
            color: AppTheme.bgElevated
            border.color: AppTheme.borderDivider
            border.width: 1
            radius: AppTheme.radiusLarge
            
            // Тень-картинка вместо вычисляемой (Level 1)
            AppShadow { level: 1 }

            ColumnLayout {
                id: newHkLayout
                anchors.top: parent.top; anchors.left: parent.left; anchors.right: parent.right
                anchors.margins: AppTheme.spaceL 
                spacing: AppTheme.spaceL 

                // ВЕРХНЯЯ СТРОКА: Название + Инпут + Выбор + Кнопка
                RowLayout {
                    Layout.fillWidth: true
                    spacing: AppTheme.spaceM

                    // ПОЛЕ НАЗВАНИЯ
                    AppTextField {
                        id: hkNameInput
                        Layout.preferredWidth: 200
                        label: "Название (для справки):"
                        placeholderText: "Напр: Ночная смена"
                    }

                    // ПОЛЕ ЗАХВАТА КЛАВИШИ
                    Rectangle {
                        Layout.preferredWidth: 200
                        Layout.preferredHeight: 44
                        radius: AppTheme.radiusMedium
                        color: keyCatcher.activeFocus ? AppTheme.bgInput : AppTheme.bgBase
                        border.color: keyCatcher.activeFocus ? AppTheme.borderFocus : AppTheme.borderInput
                        border.width: keyCatcher.activeFocus ? 2 : 1
                        
                        Text { 
                            anchors.centerIn: parent
                            text: hotkeyPageRoot.capturedKey === "" ? "Нажмите клавишу..." : hotkeyPageRoot.capturedKey
                            color: hotkeyPageRoot.capturedKey === "" ? AppTheme.textTertiary : AppTheme.accentBrand
                            font.family: AppTheme.fontFamily; font.weight: AppTheme.weightBold; font.pixelSize: AppTheme.sizeBody
                        }

                        Item {
                            id: keyCatcher
                            anchors.fill: parent
                            focus: true
                            Keys.onPressed: (event) => {
                                if (event.key === Qt.Key_Control || event.key === Qt.Key_Shift || event.key === Qt.Key_Alt) return;
                                
                                let seq = "";
                                if (event.modifiers & Qt.ControlModifier) seq += "Ctrl+";
                                if (event.modifiers & Qt.AltModifier) seq += "Alt+";
                                if (event.modifiers & Qt.ShiftModifier) seq += "Shift+";
                                
                                let keyText = event.text.toUpperCase();
                                
                                if (keyText === "") {
                                    if (event.key === Qt.Key_Space) keyText = "Space";
                                    else if (event.key === Qt.Key_Backspace) keyText = "Backspace";
                                    else if (event.key === Qt.Key_Return || event.key === Qt.Key_Enter) keyText = "Enter";
                                }
                                
                                if (keyText !== "") {
                                    hotkeyPageRoot.capturedKey = seq + keyText;
                                    event.accepted = true;
                                }
                            }
                        }
                        MouseArea { anchors.fill: parent; cursorShape: Qt.IBeamCursor; onClicked: keyCatcher.forceActiveFocus() }
                    }

                    // ВЫБОР ДЕЙСТВИЯ
                    AppComboBox {
                        id: actionTypeCombo
                        Layout.fillWidth: true
                        label: "Что сделать?"
                        model: [{text: "Добавить дежурство", value: 0}, {text: "Добавить компенсацию", value: 1}, {text: "Установить статус (Б/О/К)", value: 2}]
                        textRole: "text"; valueRole: "value"
                        onActivated: hotkeyPageRoot.currentActionTab = currentValue
                    }
                    
                    // КНОПКА СОХРАНИТЬ
                    AppButton {
                        Layout.preferredHeight: 44
                        Layout.preferredWidth: 120
                        text: "Добавить"
                        variant: "success"
                        onClicked: {
                            if (hotkeyPageRoot.capturedKey === "") { backend.showToast("Укажите клавишу!", "error"); return; }
                            
                            let newAction = { 
                                "key": hotkeyPageRoot.capturedKey,
                                "name": hkNameInput.text.trim()
                            };
                            
                            if (hotkeyPageRoot.currentActionTab === 0) {
                                newAction["type"] = "duty";
                                let sH = Math.floor(hkTimeInput.startMinutes / 60); let sM = Math.floor(hkTimeInput.startMinutes % 60);
                                let eH = Math.floor(hkTimeInput.endMinutes / 60); let eM = Math.floor(hkTimeInput.endMinutes % 60);
                                newAction["duty_start"] = ("0" + sH).slice(-2) + ":" + ("0" + sM).slice(-2);
                                newAction["duty_end"] = ("0" + eH).slice(-2) + ":" + ("0" + eM).slice(-2);
                                newAction["duty_shift"] = hkShiftCheck.checked;
                                
                                let breaksArray = [];
                                for (let i = 0; i < hotkeyPageRoot.activeBreaks.length; i++) {
                                    let b = hotkeyPageRoot.activeBreaks[i];
                                    let bsH = Math.floor(b.start / 60); let bsM = Math.floor(b.start % 60);
                                    let beH = Math.floor(b.end / 60); let beM = Math.floor(b.end % 60);
                                    breaksArray.push({ "start": ("0" + bsH).slice(-2) + ":" + ("0" + bsM).slice(-2), "end": ("0" + beH).slice(-2) + ":" + ("0" + beM).slice(-2) });
                                }
                                newAction["duty_breaks"] = breaksArray;
                                
                            } else if (hotkeyPageRoot.currentActionTab === 1) {
                                newAction["type"] = "comp";
                                newAction["comp_unit"] = hkCompUnit.currentValue;
                                newAction["comp_amount"] = hkCompUnit.currentValue === "days" ? 1 : parseInt(hkCompAmt.text);
                                
                            } else if (hotkeyPageRoot.currentActionTab === 2) {
                                newAction["type"] = "status";
                                newAction["status_val"] = hkStatusUnit.currentValue;
                            }
                            
                            let arr = [];
                            for(let i=0; i<backend.hotkeysList.length; i++) arr.push(backend.hotkeysList[i]);
                            arr.push(newAction);
                            backend.saveHotkeys(JSON.stringify(arr));
                            
                            hotkeyPageRoot.capturedKey = "";
                            hkNameInput.text = "";
                            hotkeyPageRoot.activeBreaks = [];
                        }
                    }
                }

                // НИЖНЯЯ СТРОКА: ДИНАМИЧЕСКИЕ НАСТРОЙКИ
                StackLayout {
                    Layout.fillWidth: true
                    currentIndex: hotkeyPageRoot.currentActionTab
                    
                    // ТАБ 0: ДЕЖУРСТВО
                    ColumnLayout {
                        spacing: AppTheme.spaceM
                        Layout.fillWidth: true
                        
                        RowLayout {
                            spacing: AppTheme.spaceL; Layout.fillWidth: true
                            AppTimeInterval { id: hkTimeInput; startMinutes: 480; endMinutes: 1200 }
                            AppCheckBox { id: hkShiftCheck; text: "Сменный график"; checked: true; Layout.alignment: Qt.AlignVCenter }
                            Item { Layout.fillWidth: true } 
                        }
                        
                        ColumnLayout {
                            spacing: AppTheme.spaceXS; Layout.fillWidth: true
                            Text { text: "Перерывы (опционально):"; color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall }
                            
                            Flow {
                                Layout.fillWidth: true
                                spacing: AppTheme.spaceS
                                Repeater {
                                    model: hotkeyPageRoot.activeBreaks
                                    AppPill {
                                        removable: true
                                        text: {
                                            let sH = Math.floor(modelData.start / 60); let sM = Math.floor(modelData.start % 60);
                                            let eH = Math.floor(modelData.end / 60); let eM = Math.floor(modelData.end % 60);
                                            return ("0"+sH).slice(-2) + ":" + ("0"+sM).slice(-2) + " — " + ("0"+eH).slice(-2) + ":" + ("0"+eM).slice(-2);
                                        }
                                        onRemoveClicked: { let arr = hotkeyPageRoot.activeBreaks.slice(); arr.splice(index, 1); hotkeyPageRoot.activeBreaks = arr; }
                                    }
                                }
                                AppPill { isAction: true; text: "Добавить перерыв"; onClicked: addHkBreakDialog.showAt(this, width / 2, height) }
                            }
                        }
                    }

                    // ТАБ 1: КОМПЕНСАЦИЯ
                    RowLayout {
                        spacing: AppTheme.spaceM; Layout.fillWidth: true
                        AppComboBox {
                            id: hkCompUnit; Layout.preferredWidth: 250; label: "Вид компенсации:"
                            model: [{text: "Часы (ночные)", value: "hours"}, {text: "Сверх нормы (часы)", value: "overtime"}, {text: "Дни", value: "days"}]
                            textRole: "text"; valueRole: "value"
                        }
                        AppTextField { 
                            id: hkCompAmt; Layout.preferredWidth: 150; label: "Количество:"
                            text: "8"
                            visible: hkCompUnit.currentValue !== "days" 
                        }
                        Text {
                            visible: hkCompUnit.currentValue === "days"
                            text: "Применяется ровно на 1 день"
                            color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                            Layout.alignment: Qt.AlignVCenter
                        }
                        Item { Layout.fillWidth: true }
                    }

                    // ТАБ 2: СТАТУСЫ
                    RowLayout {
                        spacing: AppTheme.spaceM; Layout.fillWidth: true
                        AppComboBox {
                            id: hkStatusUnit; Layout.preferredWidth: 250; label: "Какой статус поставить?"
                            model: [{text: "Больничный (Б)", value: "Б"}, {text: "Отпуск (О)", value: "О"}, {text: "Командировка (К)", value: "К"}]
                            textRole: "text"; valueRole: "value"
                        }
                        Item { Layout.fillWidth: true }
                    }
                }
            }
        }

        // =========================================================
        // 3. СПИСОК ДОБАВЛЕННЫХ КЛАВИШ
        // =========================================================
        Rectangle {
            Layout.fillWidth: true
            Layout.fillHeight: true
            color: "transparent"
            
            ListView {
                anchors.fill: parent; spacing: AppTheme.spaceS; clip: true; model: backend.hotkeysList
                delegate: Rectangle {
                    width: ListView.view.width; height: 56; radius: AppTheme.radiusMedium
                    color: AppTheme.bgElevated
                    border.color: AppTheme.borderDivider; border.width: 1
                    
                    RowLayout {
                        anchors.fill: parent; anchors.margins: AppTheme.spaceM; anchors.rightMargin: 60; spacing: AppTheme.spaceL
                        
                        Row {
                            spacing: AppTheme.spaceS; Layout.preferredWidth: 120
                            IconImage { source: "../icons/command.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.accentBrand; anchors.verticalCenter: parent.verticalCenter }
                            Text { text: modelData.key; color: AppTheme.accentBrand; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold; anchors.verticalCenter: parent.verticalCenter }
                        }
                        
                        Column {
                            Layout.fillWidth: true
                            spacing: 2
                            Layout.alignment: Qt.AlignVCenter

                            Text { 
                                Layout.fillWidth: true; elide: Text.ElideRight
                                color: AppTheme.textPrimary
                                font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody
                                font.weight: (modelData.name && modelData.name !== "") ? AppTheme.weightBold : AppTheme.weightMedium
                                text: {
                                    let hk = modelData
                                    // Если есть название — показываем его
                                    if (hk.name && hk.name !== "") return hk.name
                                    
                                    // Иначе — старый алгоритм
                                    if (hk.type === "duty") return "Дежурство: " + hk.duty_start + " - " + hk.duty_end
                                    if (hk.type === "status") return "Установить статус: " + hk.status_val
                                    let unt = hk.comp_unit === "hours" ? "Ночные" : (hk.comp_unit === "days" ? "Дни" : "Сверх нормы")
                                    return "Компенсация: " + unt + " (" + hk.comp_amount + ")"
                                }
                            }

                            Text {
                                visible: modelData.name && modelData.name !== ""
                                Layout.fillWidth: true; elide: Text.ElideRight
                                color: AppTheme.textSecondary
                                font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                                text: {
                                    let hk = modelData
                                    if (hk.type === "duty") return "Дежурство: " + hk.duty_start + " — " + hk.duty_end
                                    if (hk.type === "status") return "Статус: " + hk.status_val
                                    let unt = hk.comp_unit === "hours" ? "Ночные" : (hk.comp_unit === "days" ? "Дни" : "Сверх нормы")
                                    return "Компенсация: " + unt + " ×" + hk.comp_amount
                                }
                            }
                        }
                    }
                    
                    // КНОПКА УДАЛЕНИЯ
                    Rectangle {
                        width: 32; height: 32; radius: AppTheme.radiusSmall; anchors.right: parent.right; anchors.rightMargin: AppTheme.spaceS; anchors.verticalCenter: parent.verticalCenter
                        color: delKeyHover.pressed ? AppTheme.statePress : (delKeyHover.containsMouse ? AppTheme.bgDangerSoft : "transparent")
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        IconImage { anchors.centerIn: parent; source: "../icons/trash.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: delKeyHover.containsMouse ? AppTheme.accentDanger : AppTheme.textSecondary }
                        MouseArea { 
                            id: delKeyHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; 
                            onClicked: { let arr = []; for(let i=0; i<backend.hotkeysList.length; i++) arr.push(backend.hotkeysList[i]); arr.splice(index, 1); backend.saveHotkeys(JSON.stringify(arr)) } 
                        }
                    }
                }
            }
            Text { visible: backend.hotkeysList.length === 0; anchors.centerIn: parent; text: "Горячие клавиши не настроены"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody }
        }
    }

    // Всплывающее окошко для добавления перерыва
    AppDialog {
        id: addHkBreakDialog
        parent: Overlay.overlay 
        width: 320; title: ""; acceptText: "Сохранить"
        AppTimeInterval { id: newHkBreakTimeInput; anchors.horizontalCenter: parent.horizontalCenter; startMinutes: 13 * 60; endMinutes: 14 * 60 }
        onAccepted: {
            let arr = hotkeyPageRoot.activeBreaks.slice();
            arr.push({ start: newHkBreakTimeInput.startMinutes, end: newHkBreakTimeInput.endMinutes });
            arr.sort(function(a, b) { return a.start - b.start });
            hotkeyPageRoot.activeBreaks = arr;
            addHkBreakDialog.close();
        }
    }
}
