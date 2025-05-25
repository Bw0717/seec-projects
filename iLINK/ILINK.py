import sys
import json
import time
import platform
import subprocess
from PyQt6.QtGui import QFont
import requests
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QTabWidget, QLineEdit, QStatusBar, QComboBox)
from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtWidgets import QDialog

class ScanOnlyLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText("請輸入工號")
        self.setStyleSheet("""
                QLineEdit {
                    font-size: 20px;
                }
                QLineEdit::placeholder {
                    color: gray;
                }
            """)
        self.setStyleSheet("""
                QLineEdit {
                    font-size: 16px;
                    background-color: #ffffff;
                }
                QLineEdit:focus {
                    background-color: #ffffcc;
                }
            """)
        self.setReadOnly(False)  # 禁止打字
        self.last_input = ""
        self.timer = QTimer(self)
        self.timer.setInterval(300)
        self.timer.setSingleShot(True)
        #self.timer.timeout.connect(self.clear_if_too_slow) #自動清除
        self.textChanged.connect(self.on_text_changed)
        self.installEventFilter(self)
        self.setFixedSize(650, 50) 
        
    def keyPressEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.NoModifier:
            super().keyPressEvent(event)
            self.timer.start()
        else:
            event.ignore()
            
    def on_text_changed(self, text):
        self.last_input = text
        cursor_pos = self.cursorPosition()
        self.setText(text.upper())
        self.setCursorPosition(cursor_pos)


    def clear_if_too_slow(self):
        if len(self.last_input) < 5:
            self.clear()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("iLINK")
        self.resize(600, 450)
        self.api_url = "http://127.0.0.1:8000/"  #API
        self.ping_target = "127.0.0.1"            
        self.tab_locked = {"report": True}
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.status_bar = QStatusBar()
        self.status_bar.setFixedHeight(24)
        self.setStatusBar(self.status_bar)
        font = QFont("Microsoft JhengHei", 11)
        font.setBold(True)

        self.conn_status_label = QLabel("伺服器狀態：初始化中")
        self.conn_status_label.setFont(font)
        self.conn_status_label.setStyleSheet("color: black")
        
        self.action_status_label = QLabel("操作狀態：等待上工")
        self.action_status_label.setFont(font)
        self.action_status_label.setStyleSheet("color: black")

        self.status_bar.addPermanentWidget(self.conn_status_label)
        self.status_bar.addPermanentWidget(self.action_status_label)

        self.tabs.addTab(self.create_personnel_tab(), "人員上工")
        self.report_tab = self.create_report_tab()
        self.report_tab_index = self.tabs.addTab(self.report_tab, "報工作業")
        self.clear_tab = self.create_clear_tab()
        self.tabs.addTab(self.clear_tab, "清機/下工")
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.tabs.setStyleSheet("""
            QTabBar::tab {
                height: 40px;
                width: 120px;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background: lightblue;
            }
        """)
        self.start_ping_timer()

    def create_clear_tab(self):
        tab = QWidget()
        outer_layout = QHBoxLayout()
        outer_layout.setContentsMargins(10, 10, 60, 50)  
        outer_layout.setSpacing(10) 
        tab.setLayout(outer_layout)

        left_container = QWidget()
        left_container_layout = QVBoxLayout()
        left_container_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        left_container_layout.setSpacing(15)
        left_container.setLayout(left_container_layout)
        outer_layout.addWidget(left_container)
        left_container.setContentsMargins(0, 50, 0, 0) 
        
        title_label = QLabel("人員下工")
        title_label.setFixedSize(200, 50)
        title_label.setStyleSheet("font-size: 28px; font-weight: bold;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_container_layout.addWidget(title_label)

        self.logout_input = QLineEdit()
        self.logout_input.setFixedSize(200, 50)
        self.logout_input.setPlaceholderText("請輸入工號")
        self.logout_input.setStyleSheet("""
            QLineEdit {
                font-size: 20px;
                background-color: #ffffff;
            }
            QLineEdit:focus {
                background-color: #ffffcc;
            }
            QLineEdit::placeholder {
                color: gray;
            }
        """)
        left_container_layout.addWidget(self.logout_input)

        self.logout_button = QPushButton("下工")
        self.logout_button.setFixedSize(200, 50)
        self.logout_button.setStyleSheet("font-size: 20px;")
        self.logout_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.logout_button.clicked.connect(self.send_logout_request)

        left_container_layout.addWidget(self.logout_button)

        self.clear_button = QPushButton("清機")
        self.clear_button.setFixedSize(200, 50)
        self.clear_button.setStyleSheet("font-size: 20px;")
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        left_container_layout.addWidget(self.clear_button)
        self.clear_button.clicked.connect(self.send_clear_request)
        left_container_layout.addStretch()

        right_layout = QVBoxLayout()
        log_label3 = QLabel("清機/下工 Log")
        log_label3.setStyleSheet("font-size: 18px;")
        self.result_display3 = QTextEdit()
        self.result_display3.setReadOnly(True)
        self.result_display3.setMinimumWidth(400)
        self.result_display3.setMinimumHeight(350)
        self.result_display3.setStyleSheet("font-size: 14px; background-color: #f9f9f9;")

        right_layout.addWidget(log_label3)
        right_layout.addWidget(self.result_display3)
        right_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        outer_layout.addLayout(right_layout)

        return tab

    def create_personnel_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        input_layout = QHBoxLayout()
        self.scan_input = ScanOnlyLineEdit()
        self.scan_button = QPushButton("上工")
        self.scan_clear = QPushButton("清除")
        self.scan_input.returnPressed.connect(lambda: self.scan_button.setFocus())
        self.scan_clear.clicked.connect(lambda: self.scan_input.clear())
        self.scan_clear.setFixedSize(100, 50) 
        self.scan_clear.setStyleSheet("""
            QPushButton {
                font-size: 18px;
            }
        """)
        self.scan_button.clicked.connect(self.send_scan_data)
        self.scan_button.setFixedSize(100, 50) 
        self.scan_button.setStyleSheet("""
            QPushButton {
                font-size: 18px;
            }
        """)
        input_layout.addWidget(self.scan_input)
        input_layout.addWidget(self.scan_button)
        input_layout.addWidget(self.scan_clear)

        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)
        self.result_display.setFixedSize(650, 250) 
        layout.addLayout(input_layout)
        layout.addWidget(self.result_display)
        tab.setLayout(layout)
        return tab

    def create_report_tab(self):
        tab = QWidget()
        outer_layout = QHBoxLayout()
        tab.setLayout(outer_layout)

        left_layout = QVBoxLayout()
        outer_layout.addLayout(left_layout)

        program_label = QLabel("程式設定")
        font = QFont("Microsoft JhengHei", 16)
        font.setBold(True)
        program_label.setFont(font)
        left_layout.addWidget(program_label)

        combo_row = QWidget()
        combo_button_layout = QHBoxLayout()
        combo_button_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        combo_button_layout.setContentsMargins(0, 0, 0, 0)
        combo_button_layout.setSpacing(5)
        combo_row.setLayout(combo_button_layout)

        self.program_combo = QComboBox()
        self.load_program_code_map()  #JSON 載入
        self.program_combo.addItems(list(self.program_code_map.keys()))
        self.program_combo.setFixedSize(200, 45)
        self.program_combo.setStyleSheet("font-size: 16px;")
        combo_button_layout.addWidget(self.program_combo)

        self.program_button = QPushButton("上傳")
        self.program_button.setFixedSize(100, 45)
        self.program_button.setStyleSheet("font-size: 16px;")
        self.program_button.clicked.connect(self.send_program_setting)
        combo_button_layout.addWidget(self.program_button)

        self.edit_program_button = QPushButton("設定")
        self.edit_program_button.setFixedSize(100, 45)
        self.edit_program_button.setStyleSheet("font-size: 16px;")
        self.edit_program_button.clicked.connect(self.open_program_editor)
        combo_button_layout.addWidget(self.edit_program_button)

        left_layout.addWidget(combo_row)

        sn_row = QWidget()
        sn_layout = QHBoxLayout()
        sn_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        sn_layout.setContentsMargins(0, 0, 0, 0)
        sn_layout.setSpacing(5)
        sn_row.setLayout(sn_layout)

        sn_label = QLabel("SN序號")
        sn_label.setFixedSize(70, 45)
        sn_label.setStyleSheet("font-size: 16px;")
        sn_layout.addWidget(sn_label)

        self.sn_input = QLineEdit()
        self.sn_input.setFixedSize(330, 45)
        self.sn_input.setStyleSheet("font-size: 14px; background-color: yellow;")
        self.sn_input.setEnabled(False)
        self.sn_input.returnPressed.connect(self.verify_sn_input)

        sn_layout.addWidget(self.sn_input)

        left_layout.addWidget(sn_row)

        self.device_inputs = []

        for i in range(5):
            row = QWidget()
            row_layout = QHBoxLayout()
            row_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(5)
            row.setLayout(row_layout)

            field = QLineEdit()
            field.setReadOnly(False)
            field.setEnabled(False)
            field.setFixedSize(300, 40)
            field.setStyleSheet("""
                QLineEdit {
                    font-size: 16px;
                    background-color: #ffffff;
                }
                QLineEdit:focus {
                    background-color: #ffffcc;
                }
            """)
            self.device_inputs.append(field)
            row_layout.addWidget(field)

            clear_btn = QPushButton("清除")
            clear_btn.setFixedSize(70, 45)
            clear_btn.setStyleSheet("font-size: 16px;")
            clear_btn.clicked.connect(lambda _, f=field: f.clear())
            row_layout.addWidget(clear_btn)
            if i < 4:
                field.returnPressed.connect(lambda idx=i: self.device_inputs[idx + 1].setFocus())
            else:
                field.returnPressed.connect(self.confirm_and_execute_report)

            left_layout.addWidget(row)

        left_layout.addStretch()

        self.result_display2 = QTextEdit()
        self.result_display2.setReadOnly(True)
        self.result_display2.setFixedSize(400, 300)
        outer_layout.addWidget(self.result_display2)

        return tab


    def load_program_code_map(self):
        import os, json
        self.program_code_map = {}
        if os.path.exists("program_map.json"):
            with open("program_map.json", "r", encoding="utf-8") as f:
                self.program_code_map = json.load(f)

    def open_program_editor(self):
        import json, os
        dialog = QDialog(self)
        dialog.setWindowTitle("編輯程式清單")
        layout = QVBoxLayout(dialog)

        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["機種", "代號"])
        table.horizontalHeader().setStretchLastSection(True)

        self.load_program_code_map()

        for model, code in self.program_code_map.items():
            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(model))
            table.setItem(row, 1, QTableWidgetItem(code))

        layout.addWidget(table)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("新增")
        del_btn = QPushButton("刪除")
        save_btn = QPushButton("儲存並關閉")

        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(del_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

        def add_row():
            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(""))
            table.setItem(row, 1, QTableWidgetItem(""))

        def del_selected_rows():
            selected_indexes = table.selectedIndexes()
            selected_rows = set(index.row() for index in selected_indexes)
            for row in sorted(selected_rows, reverse=True):
                table.removeRow(row)

        def save_and_close():
            self.program_combo.clear()
            self.program_code_map = {}
            for row in range(table.rowCount()):
                model_item = table.item(row, 0)
                code_item = table.item(row, 1)
                if model_item and code_item:
                    model = model_item.text().strip()
                    code = code_item.text().strip().zfill(3)  # 自動補零三碼
                    if model and code:
                        self.program_combo.addItem(model)
                        self.program_code_map[model] = code

            with open("program_map.json", "w", encoding="utf-8") as f:
                json.dump(self.program_code_map, f, ensure_ascii=False, indent=2)

            dialog.accept()

        add_btn.clicked.connect(add_row)
        del_btn.clicked.connect(del_selected_rows)
        save_btn.clicked.connect(save_and_close)

        dialog.exec()

    def confirm_and_execute_report(self):
        from PyQt6.QtWidgets import QMessageBox
        reply = QMessageBox.question(self, "執行確認", "是否執行報工作業？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.execute_report()
        else:
            pass  
    def handle_sn_input_enter(self):
        for field in self.device_inputs:
            field.setEnabled(True)
        self.device_inputs[0].setFocus()
        self.device_inputs[0].selectAll()
        self.sn_input.setEnabled(False)

    def send_program_setting(self):
        import requests, json, time
        selected_model = self.program_combo.currentText()
        code = self.program_code_map.get(selected_model, "000")
        headers = {"Content-Type": "application/json"}
        payload = {
            "Content": {
                "Recipe": code.zfill(3),
                "Equipment": "AMGMESGY6003"
            },
            "FunctionName": "EquipmentAddRecipe",
            "FunctionUID": None,
            "FunctionType": "S"
        }
        try:
            start = time.time()
            self.result_display2.append("程式設定請求封包：")
            self.result_display2.append(json.dumps(payload, indent=2, ensure_ascii=False))
            res = requests.post(self.api_url+"start_program", headers=headers, data=json.dumps(payload), timeout=3)
            elapsed = int((time.time() - start) * 1000)
            res.raise_for_status()
            result = res.json()
            if result.get("ReturnCode") == "00":
                self.result_display2.append(f"\n 程式設定成功：")
                self.action_status_label.setStyleSheet("color: green")
                self.result_display2.append(json.dumps(result, indent=2, ensure_ascii=False))
                self.action_status_label.setText(f"🟢程式設定成功（{elapsed}ms）")
                self.program_button.setEnabled(False)
                self.sn_input.setEnabled(True)
            else:
                self.result_display2.append(f"\n 程式設定失敗：")
                self.result_display2.append(json.dumps(result, indent=2, ensure_ascii=False))
                self.action_status_label.setText("🔴程式設定錯誤")
                self.action_status_label.setStyleSheet("color: red")


        except Exception as e:
            self.action_status_label.setText("🔴程式設定失敗")
            self.action_status_label.setStyleSheet("color: red")
            self.result_display2.append(f"錯誤：{str(e)}")


            
    def verify_sn_input(self):
        import requests, json, time
        headers = {"Content-Type": "application/json"}
        payload = {
            "Content": {
                "Lot": self.sn_input.text().strip(),
                "Equipment": "MESARG052"
            },
            "FunctionName": "OperationVerify",
            "FunctionUID": None,
            "FunctionType": "S"
        }
        try:
            start = time.time()

            self.result_display2.append("作業確認請求封包：")
            self.result_display2.append(json.dumps(payload, indent=2, ensure_ascii=False))

            res = requests.post("http://127.0.0.1:8000/verify_operation", headers=headers, data=json.dumps(payload), timeout=3)
            elapsed = int((time.time() - start) * 1000)
            res.raise_for_status()
            result = res.json()

            if result.get("ReturnCode") == "00":
                self.result_display2.append(f"\n 作業確認成功（{elapsed}ms）：")
                self.result_display2.append(json.dumps(result, indent=2, ensure_ascii=False))
                self.action_status_label.setText(f"🟢作業確認成功（{elapsed}ms）")
                self.action_status_label.setStyleSheet("color: green")
                self.handle_sn_input_enter()
            else:
                self.result_display2.append(f"\n 作業確認失敗：")
                self.result_display2.append(json.dumps(result, indent=2, ensure_ascii=False))
                self.action_status_label.setText("🔴作業確認錯誤")
                self.action_status_label.setStyleSheet("color: red")
                self.sn_input.clear()
                self.sn_input.setFocus()

        except Exception as e:
            self.result_display2.append(f"\n 作業確認錯誤：{str(e)}")
            self.action_status_label.setText("🔴作業確認錯誤")
            self.action_status_label.setStyleSheet("color: red")
            self.sn_input.clear()
            self.sn_input.setFocus()

            
    def execute_report(self):
        import requests, json, time
        headers = {"Content-Type": "application/json"}

        dc_data = [field.text().strip() if field.text().strip() else "NA" for field in self.device_inputs]

        payload = {
            "Content": {
                "Lot": self.sn_input.text().strip(),
                "OperationResult": "00",
                "Quantity": "1",
                "MAT_SN": [],
                "DCData": dc_data,
                "Equipment": "QAA17026"
            },
            "FunctionName": "OperationMove",
            "FunctionUID": None,
            "FunctionType": "S"
        }

        try:
            start = time.time()
            self.result_display2.append("報工作業請求封包：")
            self.result_display2.append(json.dumps(payload, indent=2, ensure_ascii=False))
            res = requests.post("http://127.0.0.1:8000/execute_report", headers=headers, data=json.dumps(payload), timeout=3)
            elapsed = int((time.time() - start) * 1000)
            res.raise_for_status()
            result = res.json()
            if result.get("ReturnCode") == "00":
                self.result_display2.append(
                    f"報工成功（{elapsed}ms）\n{json.dumps(result, indent=2, ensure_ascii=False)}"
                )

                self.action_status_label.setText(f"🟢報工成功（{elapsed}ms）")
                self.action_status_label.setStyleSheet("color: green")
                self.clear_sec_all()
            else:
                self.result_display2.append(f"報工失敗：{json.dumps(result, indent=2, ensure_ascii=False)}")
                self.action_status_label.setText("🔴報工失敗")
                self.action_status_label.setStyleSheet("color: red")
        except Exception as e:
            self.result_display2.append(f"報工錯誤：{str(e)}")
            self.action_status_label.setText("🔴報工錯誤")
            self.action_status_label.setStyleSheet("color: red")


    def clear_sec_all(self):
        self.sn_input.clear()
        for field in self.device_inputs:
            field.clear()
            field.setEnabled(False)
        self.sn_input.setEnabled(True)
        self.sn_input.setFocus()

    def on_tab_changed(self, index):
        if index == self.report_tab_index and self.tab_locked["report"]:
            self.action_status_label.setText("請先完成人員上工")
            self.action_status_label.setStyleSheet("color: orange")
            self.tabs.setCurrentIndex(0)

    def send_scan_data(self):
        scan_code = self.scan_input.text().strip()
        if not scan_code:
            self.result_display.setText("錯誤：請使用Barcode Scanner掃描工號")
            self.action_status_label.setText("未輸入工號")
            self.action_status_label.setStyleSheet("color: red")
            return

        headers = {"Content-Type": "application/json"}
        payload = {
            "Content": {
                "UserID": scan_code,
                "Equipment": "iLINK123"  #設備ID
            },
            "FunctionName": "EquipmentAddUser",
            "FunctionUID": None,
            "FunctionType": "S"
        }
        try:
            start_time = time.time()
            response = requests.post(self.api_url+"start_work", headers=headers, data=json.dumps(payload), timeout=3)
            elapsed = int((time.time() - start_time) * 1000)
            response.raise_for_status()
            data = response.json()
            self.result_display.append("上工請求封包：")
            self.result_display.append(json.dumps(payload, indent=2, ensure_ascii=False))
            
            return_code = data.get("ReturnCode", "")

            if return_code == "00":
                self.tab_locked["report"] = False
                self.scan_button.setEnabled(False)   #LOCK按鈕
                self.action_status_label.setText(F"🟢上工完成（{elapsed}ms）")
                self.action_status_label.setStyleSheet("color: green")
                self.tabs.setCurrentIndex(self.report_tab_index)

            elif return_code == "01":
                self.action_status_label.setText("🔴上工失敗")
                self.action_status_label.setStyleSheet("color: red")
            else:
                self.action_status_label.setText(f"return_code：{return_code}")
                self.action_status_label.setStyleSheet("color: orange")

            self.result_display.append(f"上工回應（{elapsed}ms）：\n{json.dumps(data, indent=2, ensure_ascii=False)}")

        except Exception as e:
            self.result_display.setText(f"API 請求失敗：\n{str(e)}")
            self.action_status_label.setText("🔴上工失敗，請確認網路或格式")
            self.action_status_label.setStyleSheet("color: red")

    def start_ping_timer(self):
        self.ping_timer = QTimer()
        self.ping_timer.setInterval(5000)
        self.ping_timer.timeout.connect(self.ping_server)
        self.ping_timer.start()
        self.ping_server()


    def send_logout_request(self):
        import requests, json, time
        headers = {"Content-Type": "application/json"}
        payload = {
            "Content": {
                "UserID": self.logout_input.text().strip(),
                "Equipment": "MESARG052"
            },
            "FunctionName": "EquipmentRemoveUser",
            "FunctionUID": None,
            "FunctionType": "S"
        }
        try:
            start = time.time()
            self.result_display3.append("下工請求封包：")
            self.result_display3.append(json.dumps(payload, indent=2, ensure_ascii=False))
            res = requests.post("http://127.0.0.1:8000/logout_user", headers=headers, data=json.dumps(payload), timeout=3)
            elapsed = int((time.time() - start) * 1000)
            res.raise_for_status()
            result = res.json()
            if result.get("ReturnCode") == "00":
                self.result_display3.append(f"下工成功（{elapsed}ms）{json.dumps(result, indent=2, ensure_ascii=False)}")
                self.action_status_label.setText(f"🟢下工成功（{elapsed}ms）")
                self.action_status_label.setStyleSheet("color: green")
                self.tab_locked["report"] = True
                self.scan_button.setEnabled(True) 
                self.scan_input.clear()
            else:
                self.result_display3.append(f"下工失敗：{json.dumps(result, indent=2, ensure_ascii=False)}")
                self.action_status_label.setText("🔴下工失敗")
                self.action_status_label.setStyleSheet("color: red")
        except Exception as e:
            self.result_display3.append(f"下工錯誤：{str(e)}")
            self.action_status_label.setText("🔴下工錯誤")
            self.action_status_label.setStyleSheet("color: red")
        self.logout_input.clear()

    def send_clear_request(self):
        import requests, json, time
        headers = {"Content-Type": "application/json"}
        payload = {
            "Content": {
                "FLAG": "1",
                "Port": "",
                "Equipment": "AMGMESGY6003"
            },
            "FunctionName": "EquipmentRemoveMLot",
            "FunctionUID": None,
            "FunctionType": "S"
        }
        try:
            start = time.time()
            self.result_display3.append("清機請求封包：")
            self.result_display3.append(json.dumps(payload, indent=2, ensure_ascii=False))
            res = requests.post("http://127.0.0.1:8000/clear_equipment", headers=headers, data=json.dumps(payload), timeout=3)
            elapsed = int((time.time() - start) * 1000)
            res.raise_for_status()
            result = res.json()
            if result.get("ReturnCode") == "00":
                self.result_display3.append(f"清機成功（{elapsed}ms）{json.dumps(result, indent=2, ensure_ascii=False)}")
                self.action_status_label.setText(f"🟢清機成功（{elapsed}ms）")
                self.action_status_label.setStyleSheet("color: green")
                self.sn_input.clear()
                for field in self.device_inputs:
                    field.clear()
                    field.setEnabled(False)
                self.sn_input.setEnabled(False)
                self.program_button.setEnabled(True)
            else:
                self.result_display3.append(f"清機失敗：{json.dumps(result, indent=2, ensure_ascii=False)}")
                self.action_status_label.setText("🔴清機失敗")
                self.action_status_label.setStyleSheet("color: red")
        except Exception as e:
            self.result_display3.append(f"清機錯誤：{str(e)}")
            self.action_status_label.setText("🔴清機錯誤")
            self.action_status_label.setStyleSheet("color: red")

    def ping_server(self):
        system_os = platform.system()
        cmd = ["ping", "-n", "1", self.ping_target] if system_os == "Windows" else ["ping", "-c", "1", self.ping_target]

        try:
            output = subprocess.check_output(cmd, stderr=subprocess.STDOUT, universal_newlines=True)
            if "time=" in output:
                time_str = output.split("time=")[-1].split("ms")[0].strip()
            elif "時間=" in output:
                time_str = output.split("時間=")[-1].split("ms")[0].strip()
            else:
                time_str = "0"

            self.conn_status_label.setText(f"🟢 伺服器連線正常：{time_str} ms")
            self.conn_status_label.setStyleSheet("color: green")

        except subprocess.CalledProcessError:
            self.conn_status_label.setText("🔴 無法連線 API 伺服器")
            self.conn_status_label.setStyleSheet("color: red")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
