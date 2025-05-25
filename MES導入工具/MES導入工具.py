import sys, os, qrcode
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QStyle, QFileDialog, QListWidget, QListWidgetItem,
    QFrame, QCheckBox, QFormLayout, QGroupBox, QTableWidget, QTableWidgetItem,
    QScrollArea, QAbstractItemView, QSizePolicy, QMessageBox
)
from PyQt6.QtGui import QFont, QIcon, QDesktopServices, QColor, QKeySequence, QAction
from PyQt6.QtCore import Qt, QSize, QUrl
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont

def generate_qr_with_label(data, label, filename="qrcode.png", label_height=70, font_size=20, font_path=None):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
    qr_width, qr_height = img.size
    new_height = qr_height + label_height
    new_img = Image.new('RGB', (qr_width, new_height), color='white')
    new_img.paste(img, (0, 0))
    draw = ImageDraw.Draw(new_img)
    if font_path and os.path.exists(font_path):
        font = ImageFont.truetype(font_path, font_size)
    else:
        try:
            font = ImageFont.truetype("simhei.ttf", font_size)
        except:
            try:
                font = ImageFont.truetype("msyh.ttc", font_size)
            except:
                try:
                    font = ImageFont.truetype("wqy-microhei.ttc", font_size)
                except:
                    font = ImageFont.load_default()
    lines = label.split('\n')
    line_height = font_size + 5
    y = qr_height + (label_height - len(lines) * line_height) / 2
    for line in lines:
        text_width = draw.textlength(line, font=font)
        x = (qr_width - text_width) / 2
        draw.text((x, y), line, font=font, fill="black")
        y += line_height
    new_img.save(filename)
    print(f"已生成帶標籤的QR碼: {filename}")

def txt_transfer_excel():
    try:
        import win32com.client as win32
    except ImportError:
        print("請安裝 pywin32 模組")
        return
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        xlsm_path = os.path.join(base_path, "XX.xlsm")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(xlsm_path)
        excel.Application.Run("ImportAndConvertTxtToExcel")
        wb.Close(False)
        excel.Quit()
        del wb
        del excel
    except Exception as e:
        print(f"執行失敗: {str(e)}")

class ListItemWidget(QWidget):
    def __init__(self, text, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        self.label = QLabel(text)
        self.label.setFont(QFont("Microsoft JhengHei", 16))
        layout.addWidget(self.checkbox)
        layout.addWidget(self.label)
        layout.addStretch()

class RowWidget(QWidget):
    def __init__(self, extra_label=None, extra_editable=True, main_editable=True,
                 has_controls=True, on_plus=None, on_minus=None, update_status=None):
        super().__init__()
        self.extra_label = extra_label
        self.extra_editable = extra_editable
        self.main_editable = main_editable
        self.has_controls = has_controls
        self.on_plus = on_plus
        self.on_minus = on_minus
        self.update_status = update_status
        self.init_ui()
    
    def init_ui(self):
        layout = QHBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(0, 0, 0, 0)
        if self.has_controls:
            btn_layout = QHBoxLayout()
            btn_layout.setSpacing(5)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            self.plus_btn = QPushButton()
            plus_icon = QIcon.fromTheme("list-add")
            if plus_icon.isNull():
                self.plus_btn.setText("+")
            else:
                self.plus_btn.setIcon(plus_icon)
            self.plus_btn.setFixedSize(18, 18)
            self.plus_btn.setStyleSheet("QPushButton { background-color: green; color: white; }")
            self.plus_btn.clicked.connect(self.handle_plus)
            btn_layout.addWidget(self.plus_btn)
            self.minus_btn = QPushButton()
            minus_icon = QIcon.fromTheme("list-remove")
            if minus_icon.isNull():
                self.minus_btn.setText("-")
            else:
                self.minus_btn.setIcon(minus_icon)
            self.minus_btn.setFixedSize(18, 18)
            self.minus_btn.setStyleSheet("QPushButton { background-color: red; color: white; }")
            self.minus_btn.clicked.connect(self.handle_minus)
            btn_layout.addWidget(self.minus_btn)
            layout.addLayout(btn_layout)
        if self.extra_label is not None:
            self.extra_field = QLineEdit()
            if not self.extra_editable:
                self.extra_field.setText(self.extra_label)
            else:
                self.extra_field.setPlaceholderText(self.extra_label)
            self.extra_field.setFixedWidth(100)
            self.extra_field.setFont(QFont("Microsoft JhengHei", 14))
            self.extra_field.setReadOnly(not self.extra_editable)
            layout.addWidget(self.extra_field)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText("")
        self.edit.setFont(QFont("Microsoft JhengHei", 14))
        self.edit.setReadOnly(not self.main_editable)
        layout.addWidget(self.edit, stretch=1)
        btn_layout2 = QHBoxLayout()
        btn_layout2.setSpacing(5)
        self.copy_btn = QPushButton()
        copy_icon = QIcon.fromTheme("edit-copy")
        if copy_icon.isNull():
            copy_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)
        self.copy_btn.setIcon(copy_icon)
        self.copy_btn.setIconSize(QSize(24, 24))
        self.copy_btn.clicked.connect(self.copy_text)
        btn_layout2.addWidget(self.copy_btn)
        self.paste_btn = QPushButton()
        paste_icon = QIcon.fromTheme("edit-paste")
        if paste_icon.isNull():
            paste_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        self.paste_btn.setIcon(paste_icon)
        self.paste_btn.setIconSize(QSize(24, 24))
        self.paste_btn.clicked.connect(self.paste_text)
        btn_layout2.addWidget(self.paste_btn)
        self.clear_btn = QPushButton()
        clear_icon = QIcon.fromTheme("edit-clear")
        if clear_icon.isNull():
            clear_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DialogResetButton)
        self.clear_btn.setIcon(clear_icon)
        self.clear_btn.setIconSize(QSize(24, 24))
        self.clear_btn.clicked.connect(self.clear_text)
        btn_layout2.addWidget(self.clear_btn)
        layout.addLayout(btn_layout2)
        self.setLayout(layout)
    
    def handle_plus(self):
        if self.on_plus:
            self.on_plus(self)
    
    def handle_minus(self):
        if self.on_minus:
            self.on_minus(self)
    
    def copy_text(self):
        text = self.edit.text()
        QApplication.clipboard().setText(text)
        if self.update_status:
            self.update_status("已複製: " + text)
    
    def paste_text(self):
        text = QApplication.clipboard().text()
        self.edit.setText(text)
        if self.update_status:
            self.update_status("已貼上: " + text)
    
    def clear_text(self):
        self.edit.clear()
        if self.update_status:
            self.update_status("清除完成")

class ClipboardApp(QWidget):
    def __init__(self):
        super().__init__()
        self.row_widgets = []
        self.init_ui()
    
    def init_ui(self):
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(10)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        for i in range(10):
            if i < 7:
                row = RowWidget(extra_label="請輸入", extra_editable=True, main_editable=True,
                                has_controls=True, on_plus=self.add_row, on_minus=self.remove_row,
                                update_status=self.update_status)
            elif i == 7:
                row = RowWidget(extra_label="工單", extra_editable=False, main_editable=True,
                                has_controls=False, update_status=self.update_status)
            elif i == 8:
                row = RowWidget(extra_label="站卡1", extra_editable=False, main_editable=True,
                                has_controls=False, update_status=self.update_status)
            elif i == 9:
                row = RowWidget(extra_label="站卡2", extra_editable=False, main_editable=True,
                                has_controls=False, update_status=self.update_status)
            self.main_layout.addWidget(row)
            self.row_widgets.append(row)
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft JhengHei", 12))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.main_layout.addWidget(self.status_label)
        bottom_btn_layout = QHBoxLayout()
        bottom_btn_layout.setSpacing(10)
        self.clear_all_btn = QPushButton("全部清除")
        self.clear_all_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.clear_all_btn.clicked.connect(self.clear_all)
        bottom_btn_layout.addWidget(self.clear_all_btn)
        self.export_btn = QPushButton("匯出資料")
        self.export_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.export_btn.clicked.connect(self.export_data)
        bottom_btn_layout.addWidget(self.export_btn)
        self.main_layout.addLayout(bottom_btn_layout)
        self.setLayout(self.main_layout)
        self.setStyleSheet("""
            QWidget { background-color: #f0f0f0; }
            QLineEdit { border: 2px solid #8f8f91; border-radius: 8px; padding: 5px; background-color: #ffffff; }
            QPushButton { background-color: #4CAF50; border: none; padding: 5px; border-radius: 8px; }
            QPushButton:hover { background-color: #45a049; }
            QLabel { color: #333333; }
        """)
    
    def update_status(self, msg):
        self.status_label.setText(msg)
    
    def add_row(self, current_row):
        index = self.row_widgets.index(current_row)
        new_row = RowWidget(extra_label="請輸入", extra_editable=True, main_editable=True,
                            has_controls=True, on_plus=self.add_row, on_minus=self.remove_row,
                            update_status=self.update_status)
        self.row_widgets.insert(index + 1, new_row)
        self.main_layout.insertWidget(index + 1, new_row)
        self.update_status("已增加欄位")
    
    def remove_row(self, current_row):
        if len(self.row_widgets) <= 4:
            self.update_status("不能少於一個欄位")
            return
        index = self.row_widgets.index(current_row)
        self.row_widgets.pop(index)
        self.main_layout.removeWidget(current_row)
        current_row.deleteLater()
        self.update_status("已刪除欄位")
    
    def clear_all(self):
        for row in self.row_widgets:
            row.edit.clear()
            if hasattr(row, 'extra_field') and not row.extra_field.isReadOnly():
                row.extra_field.clear()
        self.update_status("已全部清除")
    
    def export_data(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "匯出資料", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "導入工具匯出"
                ws.append(["標籤", "內容"])
                for row in self.row_widgets:
                    extra = row.extra_field.text() if hasattr(row, 'extra_field') else ""
                    main = row.edit.text()
                    ws.append([extra, main])
                wb.save(file_path)
                self.update_status("匯出路徑 " + file_path)
            except Exception as e:
                self.update_status("Export failed: " + str(e))

class SubTabWidget(QWidget):
    def __init__(self, clipboard_app, parent=None):
        super().__init__(parent)
        self.clipboard_app = clipboard_app
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        top_btn_layout = QHBoxLayout()
        top_btn_layout.setSpacing(10)
        self.search_btn = QPushButton("搜尋")
        self.search_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.search_btn.clicked.connect(self.search_data)
        top_btn_layout.addWidget(self.search_btn)
        self.generate_btn = QPushButton("生成")
        self.generate_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.generate_btn.clicked.connect(self.generate_data)
        top_btn_layout.addWidget(self.generate_btn)
        self.open_folder_btn = QPushButton("Open folder")
        self.open_folder_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.open_folder_btn.clicked.connect(self.open_folder)
        top_btn_layout.addWidget(self.open_folder_btn)
        layout.addLayout(top_btn_layout)
        self.frame = QFrame()
        self.frame.setFrameShape(QFrame.Shape.StyledPanel)
        frame_layout = QVBoxLayout()
        self.list_widget = QListWidget()
        self.list_widget.setFont(QFont("Microsoft JhengHei", 16))
        frame_layout.addWidget(self.list_widget)
        self.frame.setLayout(frame_layout)
        layout.addWidget(self.frame)
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft JhengHei", 12))
        layout.addWidget(self.status_label)
        self.setLayout(layout)
    
    def open_folder(self):
        folder = os.path.abspath("QRCode_Output")
        if not os.path.exists(folder):
            os.makedirs(folder)
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
    
    def search_data(self):
        self.list_widget.clear()
        for row in self.clipboard_app.row_widgets:
            extra = row.extra_field.text() if hasattr(row, 'extra_field') else ""
            main = row.edit.text()
            item_text = f"{extra}, {main}"
            item = QListWidgetItem()
            widget = ListItemWidget(item_text)
            widget.setMinimumHeight(40)
            item.setSizeHint(widget.sizeHint())
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, widget)
    
    def generate_data(self):
        count = self.list_widget.count()
        if count == 0:
            self.status_label.setText("無資料可生成")
            self.clipboard_app.update_status("無資料可生成")
            return
        output_folder = "QRCode_Output"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        generated = False
        for index in range(count):
            item = self.list_widget.item(index)
            widget = self.list_widget.itemWidget(item)
            if widget and widget.checkbox.isChecked():
                text = widget.label.text()
                parts = text.split(',', 1)
                extra = parts[0].strip() if len(parts) > 0 else ""
                main = parts[1].strip() if len(parts) > 1 else ""
                if main == "":
                    continue
                if extra != "":
                    filename = os.path.join(output_folder, f"{extra}.png")
                else:
                    filename = os.path.join(output_folder, f"data_{index}.png")
                label = f"{main}\n{extra}" if extra else main
                generate_qr_with_label(
                    data=main,
                    label=label,
                    filename=filename,
                    label_height=70,
                    font_size=20,
                    font_path=None
                )
                generated = True
        if generated:
            self.status_label.setText("生成成功")
            self.clipboard_app.update_status("生成成功")
        else:
            self.status_label.setText("無資料生成")
            self.clipboard_app.update_status("無資料可生成")

class MaterialImporter(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlight_items = []
        self.current_match_index = -1
        self.matches = []
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("輸入關鍵字搜尋...")
        self.search_input.setFont(QFont("Microsoft JhengHei", 12))
        self.search_input.returnPressed.connect(self.handle_search)
        self.search_btn = QPushButton("搜尋")
        self.search_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.search_btn.clicked.connect(self.handle_search)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_btn)
        self.table = QTableWidget()
        self.table.setFont(QFont("Microsoft JhengHei", 12))
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table.setStyleSheet("""
            QTableView {
                selection-background-color: #FFA500;
                selection-color: black;
            }
        """)   
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft JhengHei", 12))     

        btn_layout = QHBoxLayout()
        self.import_btn = QPushButton("物料匯入")
        self.import_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.import_btn.clicked.connect(self.import_material)
        self.open_excel_btn = QPushButton("開啟Excel")
        self.open_excel_btn.setFont(QFont("Microsoft JhengHei", 12))
        self.open_excel_btn.clicked.connect(self.open_excel)
        btn_layout.addWidget(self.import_btn)
        btn_layout.addWidget(self.open_excel_btn)
        
        self.table = QTableWidget()
        self.table.setFont(QFont("Microsoft JhengHei", 12))
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft JhengHei", 12))
        
        layout.addLayout(search_layout)
        layout.addLayout(btn_layout)
        layout.addWidget(self.table, 1)
        layout.addWidget(self.status_label)
        self.setLayout(layout)

        self.search_shortcut = QAction(self)
        self.search_shortcut.setShortcut(QKeySequence("Ctrl+F"))
        self.search_shortcut.triggered.connect(self.focus_search_input)
        self.addAction(self.search_shortcut)

    def focus_search_input(self):
        if hasattr(self, 'search_input') and self.search_input:
            self.search_input.setFocus()

    def handle_search(self):
        keyword = self.search_input.text().strip().lower()
        if not hasattr(self, 'last_keyword') or self.last_keyword != keyword:
            self.current_match_index = -1
            self.last_keyword = keyword
            self.do_search(keyword)
        else:
            self.jump_to_next_match()

    def do_search(self, keyword):
        for item in self.highlight_items:
            item.setBackground(QColor(Qt.GlobalColor.white))
        self.highlight_items = []
        self.matches = []
        if not keyword:
            return
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and keyword in item.text().lower():
                    self.matches.append((row, col))
        if self.matches:
            self.highlight_all_matches()
            self.jump_to_next_match()

    def highlight_all_matches(self):
        for row, col in self.matches:
            item = self.table.item(row, col)
            item.setBackground(QColor(Qt.GlobalColor.yellow))
            self.highlight_items.append(item)

    def jump_to_next_match(self):
        try:
            if not self.matches:
                QMessageBox.information(self, "提示", "沒有匹配")
                return
            
            self.current_match_index += 1
            if self.current_match_index >= len(self.matches):
                self.current_match_index = 0
                QApplication.beep()
            
            row, col = self.matches[self.current_match_index]
            self.table.scrollToItem(self.table.item(row, col), QAbstractItemView.ScrollHint.PositionAtCenter)
            self.table.clearSelection()
            self.table.setCurrentCell(row, col)
        except Exception as e:
            print(f"錯誤: {str(e)}")
            self.status_label.setText("搜索失敗")

    def clear_highlights(self):
        for item in self.highlight_items:
            item.setBackground(QColor(Qt.GlobalColor.white))
        self.highlight_items = []
        self.matches = []
        self.current_match_index = -1

    def search_table(self):
        keyword = self.search_input.text().strip().lower()

        for item in self.highlight_items:
            item.setBackground(QColor(Qt.GlobalColor.white))
        self.highlight_items = []
        
        if not keyword:
            return

        found = False
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and keyword in item.text().lower():
                    item.setBackground(QColor(Qt.GlobalColor.yellow))
                    self.highlight_items.append(item)
                    if not found:
                        self.table.scrollToItem(item, QAbstractItemView.ScrollHint.PositionAtTop)
                        found = True
    
    def import_material(self):
        txt_transfer_excel()
        print("物料匯入完成。")
        self.status_label.setText("物料匯入完成")
    
    def open_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path and file_path != "":
            self.load_excel(file_path)
            self.status_label.setText("載入 " + file_path)
        else:
            print("取消匯入")
            self.status_label.setText("取消匯入")
    
    def load_excel(self, file_path):
        try:
            df = pd.read_excel(file_path)
            rows, cols = df.shape
            self.table.setRowCount(rows)
            self.table.setColumnCount(cols)
            self.table.setHorizontalHeaderLabels(df.columns)
            for i in range(rows):
                for j in range(cols):
                    value = df.iloc[i, j]
                    item = QTableWidgetItem(str(value))
                    self.table.setItem(i, j, item)
            self.table.resizeColumnsToContents()
        except Exception as e:
            print("Error loading Excel file:", e)
            self.status_label.setText("Error loading Excel file: " + str(e))

class WeightCalculator(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle('包裝重量計算工具')
        font_label = QFont("Microsoft JhengHei", 16)
        font_edit = QFont("Microsoft JhengHei", 16)
        font_button = QFont("Microsoft JhengHei", 16)
        
        input_group = QGroupBox('輸入')
        input_group.setFont(font_label)
        input_layout = QFormLayout()
        input_layout.setContentsMargins(5, 5, 5, 5)
        input_layout.setSpacing(10)
        
        self.single_weight_label = QLabel('單顆重量：')
        self.single_weight_label.setFont(font_label)
        self.single_weight_edit = QLineEdit()
        self.single_weight_edit.setFont(font_edit)
        self.single_weight_edit.setMinimumHeight(30)
        self.single_weight_edit.textChanged.connect(self.calculate_limits)
        input_layout.addRow(self.single_weight_label, self.single_weight_edit)
        
        self.quantity_per_box_label = QLabel('每箱數量：')
        self.quantity_per_box_label.setFont(font_label)
        self.quantity_per_box_edit = QLineEdit()
        self.quantity_per_box_edit.setFont(font_edit)
        self.quantity_per_box_edit.setMinimumHeight(30)
        self.quantity_per_box_edit.textChanged.connect(self.calculate_limits)
        input_layout.addRow(self.quantity_per_box_label, self.quantity_per_box_edit)
        
        self.packaging_weight_label = QLabel('包材重量：')
        self.packaging_weight_label.setFont(font_label)
        self.packaging_weight_edit = QLineEdit()
        self.packaging_weight_edit.setFont(font_edit)
        self.packaging_weight_edit.setMinimumHeight(30)
        self.packaging_weight_edit.textChanged.connect(self.calculate_limits)
        input_layout.addRow(self.packaging_weight_label, self.packaging_weight_edit)
        
        input_group.setLayout(input_layout)
        
        output_group = QGroupBox('輸出')
        output_group.setFont(font_label)
        output_layout = QFormLayout()
        output_layout.setContentsMargins(5, 5, 5, 5)
        output_layout.setSpacing(10)
        
        self.lower_limit_label = QLabel('下限值：')
        self.lower_limit_label.setFont(font_label)
        self.lower_limit_display = QLineEdit()
        self.lower_limit_display.setFont(font_edit)
        self.lower_limit_display.setMinimumHeight(30)
        self.lower_limit_display.setReadOnly(True)
        output_layout.addRow(self.lower_limit_label, self.lower_limit_display)
        
        self.upper_limit_label = QLabel('上限值：')
        self.upper_limit_label.setFont(font_label)
        self.upper_limit_display = QLineEdit()
        self.upper_limit_display.setFont(font_edit)
        self.upper_limit_display.setMinimumHeight(30)
        self.upper_limit_display.setReadOnly(True)
        output_layout.addRow(self.upper_limit_label, self.upper_limit_display)
        
        self.clear_button = QPushButton('清除')
        self.clear_button.setFont(font_button)
        self.clear_button.setMinimumHeight(35)
        self.clear_button.clicked.connect(self.clear_inputs)
        
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        button_layout.addStretch(1)
        button_layout.addWidget(self.clear_button)
        output_layout.addRow(button_layout)
        
        output_group.setLayout(output_layout)
        
        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.addWidget(input_group)
        main_layout.addWidget(output_group)
        self.setLayout(main_layout)
    
    def calculate_limits(self):
        try:
            single_weight = float(self.single_weight_edit.text())
            quantity_per_box = int(self.quantity_per_box_edit.text())
            packaging_weight = float(self.packaging_weight_edit.text())
            total_weight = single_weight * quantity_per_box 
            tolerance = total_weight * 0.1
            upper_limit = round(total_weight + tolerance + packaging_weight, 2)
            lower_limit = round(total_weight - tolerance + packaging_weight, 2)
            self.upper_limit_display.setText(f'{upper_limit:.2f}')
            self.lower_limit_display.setText(f'{lower_limit:.2f}')
        except ValueError:
            self.upper_limit_display.setText('輸入無效')
            self.lower_limit_display.setText('輸入無效')
    
    def clear_inputs(self):
        self.single_weight_edit.clear()
        self.quantity_per_box_edit.clear()
        self.packaging_weight_edit.clear()
        self.upper_limit_display.clear()
        self.lower_limit_display.clear()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("導入工具")
        self.resize(500, 650)
        self.setMinimumSize(500, 650)
        self.init_ui()
    
    def init_ui(self):
        tab_widget = QTabWidget()
        
        clipboard_scroll = QScrollArea()
        clipboard_scroll.setWidgetResizable(True)
        self.clipboard_app = ClipboardApp()
        clipboard_scroll.setWidget(self.clipboard_app)
        tab_widget.addTab(clipboard_scroll, "剪貼簿")

        material_scroll = QScrollArea()
        material_scroll.setWidgetResizable(True)
        self.material_importer = MaterialImporter()
        material_scroll.setWidget(self.material_importer)
        tab_widget.addTab(material_scroll, "物料匯入")
        
        qrcode_scroll = QScrollArea()
        qrcode_scroll.setWidgetResizable(True)
        self.sub_tab = SubTabWidget(self.clipboard_app)
        qrcode_scroll.setWidget(self.sub_tab)
        tab_widget.addTab(qrcode_scroll, "QRcode生成工具")
        
        self.weight_calc = WeightCalculator()
        tab_widget.addTab(self.weight_calc, "包裝重量計算工具")
        
        self.setCentralWidget(tab_widget)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())