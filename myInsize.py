from PyQt6 import QtCore, QtWidgets # type: ignore
from PyQt6.QtWidgets import QApplication, QTableWidget, QMenu, QTableWidgetItem, QLineEdit, QSpinBox, QPushButton, QLabel, QWidget # type: ignore
from PyQt6.QtGui import QIcon, QKeySequence, QColor, QFont, QBrush # type: ignore
from PyQt6.QtCore import Qt, pyqtSignal # type: ignore
import pandas as pd # type: ignore
import sys
import os
import io
import datetime
import locale
import math
import json
import random
import subprocess
# from library import ColorPicker


encodeUTF8 = {"encoding": 'utf-8'}
# sys.stdout = io.TextIOWrapper(open(sys.stdout.fileno(), 'wb'), encoding='utf-8')
startupinfo = subprocess.STARTUPINFO()
startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

locale.setlocale(locale.LC_TIME, 'vi_VN')
now = datetime.datetime.now()
date = now.strftime('%m/%y')

if len(sys.argv) > 1: 
    input_file = sys.argv[1]
else:
    input_file = r"F:\testTem\input\SP_HANG_HOA - Copy.xlsx"

if getattr(sys, 'frozen', False):
    dirname = os.path.dirname(sys.executable)
else:
    dirname = os.path.dirname(os.path.abspath(__file__))

base_name =  os.path.basename(input_file)
file_name = base_name.replace(".xlsx", '')
outputFile = f'{dirname}\\output\\{file_name}_output.xlsx'
btwFile = f'{dirname}\\btw\\{file_name}.btw'
jsPath = f'{dirname}\\json\\batender.json'
settingPath = f'{dirname}\\json\\design.json'
iconPath = f'{dirname}\\icon\\print.ico'
cmd = f'bartend.exe /F="{btwFile}" /P /X'
os.system('chcp 65001')
main_button = [780, 480, 90, 27]
history_n = 50
        
def ceil(vl: float):
    try:
        vl = math.ceil(vl)
    except ValueError:
        pass
    return vl

def encode_value(text: str, localvalue: list):
    try:
        try:
            try:
                return str(eval(f'ceil({text})', {}, localvalue))
            except (ValueError, SyntaxError, NameError, TypeError, ZeroDivisionError):
                return str(eval(text, {}, localvalue))
        except (ValueError, SyntaxError, NameError, TypeError, ZeroDivisionError):
            return str(eval(f'f"""{text}"""', {}, localvalue))
    except (ValueError, SyntaxError, NameError, TypeError, ZeroDivisionError):
        return text

rmb = {
    'ceil': ceil,
    'date': date,
    'math': math,
    'sizelist': ('S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL', '6XL', '7XL', '8XL', '9XL', '100', '110', '120', '130', '140', '150', '160', 'SIZE', '(...)'),
    'i': 0
}

rmb['fixdict'] = {val: i for i, val in enumerate(rmb['sizelist'])}

def import_json(filepath: str) -> dict:
    with open(filepath, 'r', **encodeUTF8) as file:
        file_text = file.read()
        rmb_js = json.loads(file_text)
        return rmb_js

def save_json(filepath: str, rmb_js: dict):
    with open(filepath, 'w', **encodeUTF8) as file:
        text = json.dumps(rmb_js, sort_keys=True, indent=4)
        file.write(text)

rmb_setting = import_json(settingPath)

mainCss = f"""
    QWidget {{
        background-color: {rmb_setting['main']['background-color']};
        color: {rmb_setting['main']['text-color']};
        font-size: {rmb_setting['main']['font-size']};
        font-family: {rmb_setting['main']['font-name']};
        padding: 0;
        margin: 0;
    }}
    
    QPushButton {{
        background-color: #e0e0e0;
        border: 1px solid #bbb;
        border-radius: 13px;
    }}
    
    QPushButton:hover {{
        background-color: #33000000;
        color: black;
        font-size: 12px;
        font-weight: bold;
        border: 0;
    }}
    
    QTableWidget {{
        border: 1px solid #33000000;
        alternate-background-color: #fbfbfb;
        font-size: 12px;
        background-color: #ffffff;
        padding: 0;
        margin: 0;
    }}
    
    QTableWidget::item:selected {{
        background-color: #440088ff;
        color: black;
        font-weight: bolder;
    }}
    
    QHeaderView::section {{
        background-color: #dddddd;
        color: black;
        border: 0;
    }}
    
    QHeaderView::section:hover {{
        background-color: #33000000;
        color: black;
        font-weight: bold;
    }}
    
    QHeaderView::section:checked {{
        background-color: #330088ff;
        color: black;
        font-weight: bold;
    }}
"""

class SettingWindow(QWidget):
    colorPicked = pyqtSignal(dict)
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Setting")
        self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.WindowCloseButtonHint)
        self.setStyleSheet(mainCss)
        
        self.backgroupLabel = QLabel(parent=self)
        self.backgroupLabel.setGeometry(QtCore.QRect(10, 10, 200, 31))
        self.backgroupLabel.setObjectName("backgroupLabel")
        self.backgroupLabel.setText("Backgroup color")

class CustomTableWidget(QTableWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.rawValue = []
        self.undo_stack = []
        self.redo_stack = []
        self.undo_stacking = False
        self.newColumn_text = str('')

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Cut):
            self.cut()
        elif event.matches(QKeySequence.StandardKey.Copy):
            self.copy()
        elif event.matches(QKeySequence.StandardKey.Paste):
            self.paste()
        elif event.matches(QKeySequence.StandardKey.Undo):
            self.undo_stacking = True
            self.undo(self.undo_stack, self.redo_stack)
        elif event.matches(QKeySequence.StandardKey.Redo):
            self.undo_stacking = True
            self.undo(self.redo_stack, self.undo_stack)
        elif event.matches(QKeySequence.StandardKey.SelectAll):
            self.selectAll()
        elif event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_Left:
                self.ctrl_left_event()
            elif event.key() == Qt.Key.Key_Right:
                self.ctrl_right_event()
            elif event.key() == Qt.Key.Key_Up:
                self.add_row(sp=1)
            elif event.key() == Qt.Key.Key_Down:
                self.add_row()
            elif event.key() == Qt.Key.Key_Delete:
                self.remove_row()
        elif event.key() == Qt.Key.Key_Delete:
                self.delete_event()
        else:
            super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        context_menu = QMenu(self)
        context_menu.setStyleSheet("""
            QMenu::item {
                padding-right: 10px; /* Adjust this value to your preference */
                margin: 3px;
            }
            QMenu::item::indicator {
                padding-left: 15px; /* Adjust to properly align text */
            }
            QMenu::item:hover {
                color: #0000f0;
                font-weight: bold;
                background-color: #f0f0f0;
            }
        """)
        
        undo_action = context_menu.addAction("Quay lại")
        undo_action.setShortcut(QKeySequence("Ctrl+Z"))
        
        redo_action = context_menu.addAction("Tiến tới")
        redo_action.setShortcut(QKeySequence("Ctrl+Y"))
        
        context_menu.addSeparator()
        
        cut_action = context_menu.addAction("Cắt")
        cut_action.setShortcut(QKeySequence.StandardKey.Cut)
        
        copy_action = context_menu.addAction("Sao chép")
        copy_action.setShortcut(QKeySequence.StandardKey.Copy)

        paste_action = context_menu.addAction("Dán")
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        
        context_menu.addSeparator()
        
        add_top_row_action = context_menu.addAction("Thêm dòng trên")
        add_top_row_action.setShortcut(QKeySequence("Ctrl+Up"))
        
        add_bot_row_action = context_menu.addAction("Thêm dòng dưới")
        add_bot_row_action.setShortcut(QKeySequence("Ctrl+Down"))
        
        add_left_col_action = context_menu.addAction("Thêm cột trái")
        add_left_col_action.setShortcut(QKeySequence("Ctrl+Left"))
        
        add_right_col_action = context_menu.addAction("Thêm cột phải")
        add_right_col_action.setShortcut(QKeySequence("Ctrl+Right"))
        
        context_menu.addSeparator()
        
        delete_row_action = context_menu.addAction("Xóa dòng")
        delete_row_action.setShortcut(QKeySequence("Ctrl+Delete"))
        
        delete_column_action = context_menu.addAction("Xóa cột")
        
        delete_text_action = context_menu.addAction("Chỉ xóa nội dung")
        delete_text_action.setShortcut(QKeySequence("Delete"))
        
        context_menu.addSeparator()
        
        selecAll_action = context_menu.addAction("Chọn tất cả")
        selecAll_action.setShortcut(QKeySequence("Ctrl+A"))
        
        action = context_menu.exec(self.mapToGlobal(event.pos()))
        
        if action == undo_action:
            self.undo_stacking = True
            self.undo(self.undo_stack, self.redo_stack)
        elif action == redo_action:
            self.undo_stacking = True
            self.undo(self.redo_stack, self.undo_stack)
        elif action == cut_action:
            self.cut()
        elif action == copy_action:
            self.copy()
        elif action == paste_action:
            self.paste()
        elif action == add_top_row_action:
            self.add_row(sp=1)
        elif action == add_bot_row_action:
            self.add_row()
        elif action == add_left_col_action:
            self.ctrl_left_event()
        elif action == add_right_col_action:
            self.ctrl_right_event()
        elif action == delete_row_action:
            self.remove_row()
        elif action == delete_column_action:
            self.remove_column()
        elif action == delete_text_action:
            self.delete_event()
        elif action == selecAll_action:
            self.selectAll()
        
    def delete_event(self):
        select_items = self.selectedItems()
        for item in select_items:
            row = item.row()
            col = item.column()
            val = self.rawValue[row][col]
            self.undo_stack.append({
                'row': row,
                'col': col,
                'style': 'change_item',
                'val': val,
                'n': len(select_items)
            })
            self.setItem(row, col, QTableWidgetItem(""))
            self.rawValue[row][col] = ""
            
    def ctrl_left_event(self):
        select_items = self.selectedItems()
        if select_items:
            self.add_column(id=select_items[0].column())
            
    def ctrl_right_event(self):
        select_items = self.selectedItems()
        if select_items:
            self.add_column(id=select_items[0].column() + 1)
        
    def copy(self):
        selected_range = self.selectedRanges()[0]
        rows = selected_range.rowCount()
        cols = selected_range.columnCount()

        data = []
        for i in range(rows):
            row_data = []
            
            for j in range(cols):
                if self.rawValue[selected_range.topRow() + i]:
                    item = self.rawValue[selected_range.topRow() + i][selected_range.leftColumn() + j]
                    row_data.append(str(item) or '')
                else:
                    row_data.append('')
            data.append('\t'.join(row_data))

        clipboard = QApplication.clipboard()
        clipboard.setText('\n'.join(data))
        
    def cut(self):
        selected_range = self.selectedRanges()[0]
        rows = selected_range.rowCount()
        cols = selected_range.columnCount()

        data = []
        save_item = []
        for i in range(rows):
            row_data = []
            for j in range(cols):
                top_row = selected_range.topRow() + i
                left_col = selected_range.leftColumn() + j
                if self.rawValue[top_row]:
                    item = self.rawValue[top_row][left_col]
                    self.setItem(top_row, left_col, QTableWidgetItem(""))
                    self.rawValue[top_row][left_col] = ""
                    row_data.append(str(item) or '')
                    save_item.append({
                        'row': top_row,
                        'col': left_col,
                        'style': 'change_item',
                        'val': item
                    })
                else:
                    row_data.append('')
            data.append('\t'.join(row_data))
        for i, v in enumerate(save_item):
            save_item[i]['n'] = len(save_item)
            self.undo_stack.append(save_item[i])
        clipboard = QApplication.clipboard()
        clipboard.setText('\n'.join(data))

    def paste(self):
        clipboard = QApplication.clipboard()
        data = clipboard.text().split('\n')

        selected_range = self.selectedRanges()[0]
        row = selected_range.topRow()
        col = selected_range.leftColumn()
        paste_item = []
        for i, line in enumerate(data):
            row_data = line.split('\t')
            for j, cell in enumerate(row_data):
                paste_item.append({
                    'row': row + i,
                    'col': col + j,
                    'cell': self.rawValue[row + i][col + j]
                })
                self.setItem(row + i, col + j, QTableWidgetItem(encode_value(cell, rmb)))
                self.rawValue[row + i][col + j] = cell
        
        self.undo_stack.append({
            'style': 'paste',
            'val': paste_item
        })
        
    def undo(self, change: list, rechange: list):
        if change and len(change) > 0:
            last_state = change[-1]
            if last_state['style'] == 'change_item':
                for i in range(last_state['n']):
                    item = change.pop()
                    reRaw = self.rawValue[item['row']][item['col']]
                    rechange.append({
                        'row': item['row'],
                        'col': item['col'],
                        'style': item['style'],
                        'val': reRaw,
                        'n': item['n']
                    })
                    cell = item['val']
                    self.setItem(item['row'], item['col'], QTableWidgetItem(encode_value(cell, rmb)))
                    self.rawValue[item['row']][item['col']] = cell
            elif last_state['style'] == 'add_col':
                change.pop()
                oldval = self.remove_column(check=True, ids=last_state['col'])
                rechange.append(oldval)
            elif last_state['style'] == 'del_col':
                change.pop()
                cols = []
                for i, col in enumerate(last_state['col']):
                    col_name = last_state['colname'][i]
                    col_val = last_state['val'][i]
                    cols.append(col)
                    self.add_column(col_val, col, col_name, True)
                rechange.append({'style': 'add_col', 'col': cols})
            elif last_state['style'] == 'add_row':
                change.pop()
                nval = self.remove_row(last_state['val'], True)
                rechange.append({'style': 'del_row', 'val': nval})
            elif last_state['style'] == 'del_row':
                change.pop()
                nval = self.add_row(last_state['val'], True)
                rechange.append({'style': 'add_row', 'val': nval})
            elif last_state['style'] == 'paste':
                change.pop()
                nval = []
                for item in last_state['val']:
                    row = item['row']
                    col = item['col']
                    cell = item['cell']
                    nval.append({
                        'row': row,
                        'col': col,
                        'cell': self.rawValue[row][col]
                    })
                    self.setItem(row, col, QTableWidgetItem(encode_value(cell, rmb)))
                    self.rawValue[row][col] = cell
                rechange.append({'style': 'paste', 'val': nval})
        self.undo_stacking = False
                
    def add_column(self, val: list|tuple = None, id=0, colname="", check=False):
        column_count = id if id else self.columnCount()
        column_name = colname if colname else self.newColumn_text
        self.insertColumn(column_count)
        self.setHorizontalHeaderItem(column_count, QTableWidgetItem(column_name))
        for row in range(self.rowCount()):
            nval = str(val[row]) if val else ""
            self.setItem(row, column_count, QTableWidgetItem(encode_value(nval, rmb)))
            self.rawValue[row].insert(column_count, nval)
        if not check:
            self.undo_stack.append({
                'style': 'add_col',
                'col': [column_count]
            })
            
    def remove_column(self, check=False, ids: list|tuple = None):
        if not check:
            selected_columns = self.selectionModel().selectedColumns()
            column_index = ids if ids else [cl.column() for cl in selected_columns]
            headers = [self.horizontalHeaderItem(col).text() for col in column_index if self.horizontalHeaderItem(col)]
            rmv_col_item = [[row[id] for row in self.rawValue] for id in column_index]
            for col in sorted(column_index, reverse=True):
                self.removeColumn(col)
                for row in range(self.rowCount()):
                    del self.rawValue[row][col]
            self.undo_stack.append({
                'style': 'del_col',
                'col': column_index,
                'colname': headers,
                'val': rmv_col_item
            })
        else:
            index = ids[0] if ids else self.columnCount() - 1
            rmv_col_item = {
                'style': 'del_col',
                'col': [index],
                'colname': [self.horizontalHeaderItem(index).text()],
                'val': [[row[index] for row in self.rawValue]]
            }
            self.removeColumn(index)
            for row in range(self.rowCount()):
                del self.rawValue[row][index]
            return rmv_col_item
            
    def add_row(self, val: list|tuple = None, check=False, sp=0):
        new_row = []
        if val:
            for value in val[::-1]:
                row = value['row']
                cell = value['val']
                self.insertRow(row)
                self.rawValue.insert(row, cell)
                new_row.append(row)
                for col in range(self.columnCount()):
                    self.setItem(row, col, QTableWidgetItem(cell[col]))
        else:
            selected_items = self.selectedItems()
            new_row_index = (sp == 1 and selected_items[0].row() or selected_items[0].row() + 1) if selected_items else 0
            new_row.append(new_row_index)
            self.insertRow(new_row_index)
            self.rawValue.insert(new_row_index, ['' for col in range(self.columnCount())])
            for col in range(self.columnCount()):
                self.setItem(new_row_index, col, QTableWidgetItem(""))
                
        if not check:
            self.undo_stack.append({
                'style': 'add_row',
                'val': new_row
            })
        else:
            return new_row
        
    def remove_row(self, val: list|tuple=None, check=False):
        selected_rows = val if val else [index.row() for index in self.selectedItems()]
        selected_rows = set(selected_rows)
        rmv_row = []
        for row in sorted(selected_rows, reverse=True):
            if not self.isRowHidden(row):
                rmv_row.append({ 
                    'row': row, 
                    'val': self.rawValue[row] 
                }) 
                self.removeRow(row)
                del self.rawValue[row]
        if not check:
            self.undo_stack.append({
                'style': 'del_row',
                'val': rmv_row
            })
        else:
            return rmv_row
                                
class Ui_PrintTemp(object):
    def setupUi(self, PrintTemp):
        self.PrintTemp = PrintTemp
        self.PrintTemp.setObjectName("PrintTemp")
        self.PrintTemp.resize(1390, 754)
        self.PrintTemp.setStyleSheet(mainCss)
        
        self.rmb = import_json(jsPath)
        self.set_undo = []
        self.set_end_item = False
        
        self.label = QLabel(parent=self.PrintTemp)
        self.label.setGeometry(QtCore.QRect(20, 10, 70, 31))
        self.label.setObjectName("label")
        
        self.inputSearch = QLineEdit(parent=self.PrintTemp)
        self.inputSearch.setGeometry(QtCore.QRect(95, 11, 320, 28))
        self.inputSearch.setObjectName("inputSearch")
        self.inputSearch.setPlaceholderText("Nhập tên cần tìm...")
        self.inputSearch.textChanged.connect(self.search)
        
        self.tableWidget = CustomTableWidget(parent=self.PrintTemp)
        self.tableWidget.setGeometry(QtCore.QRect(10, 50, 1061, 425))
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableWidget.itemSelectionChanged.connect(lambda: self.on_selection_change(self.tableWidget, True))
        header = self.tableWidget.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(True)
        
        self.tableWidget_history = QTableWidget(parent=self.PrintTemp)
        self.tableWidget_history.setGeometry(QtCore.QRect(10, 515, 1370, 230))
        self.tableWidget_history.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tableWidget_history.setObjectName("tableWidget_history")
        self.tableWidget_history.setColumnCount(0)
        self.tableWidget_history.setRowCount(0)
        self.tableWidget_history.setAlternatingRowColors(True)
        self.tableWidget_history.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableWidget_history.itemSelectionChanged.connect(lambda: self.on_selection_change(self.tableWidget_history))
        header_history = self.tableWidget_history.horizontalHeader()
        header_history.setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.Stretch)
        header_history.setStretchLastSection(True)
        
        self.label_edit = QLabel(parent=self.PrintTemp)
        self.label_edit.setGeometry(QtCore.QRect(430, 10, 75, 31))
        self.label_edit.setObjectName("label_edit")
        
        self.inputCode = QSpinBox(parent=self.PrintTemp)
        self.inputCode.setGeometry(QtCore.QRect(692, 11, 70, 30))
        self.inputCode.setObjectName("inputCode")
        self.inputCode.setPrefix(" #")
        self.inputCode.setRange(0, 9999)
        self.inputCode.textChanged.connect(lambda code: self.for_seclect_items(lambda i, text, _: [text, code]))
        
        self.commandLinkButton = QPushButton(parent=self.PrintTemp)
        self.commandLinkButton.setGeometry(QtCore.QRect(970, 8, 90, 31))
        self.commandLinkButton.setObjectName("commandLinkButton")
        self.commandLinkButton.clicked.connect(lambda: subprocess.run(f'bartend.exe "{btwFile}"', startupinfo=startupinfo, shell=True))
        self.commandLinkButton.setStyleSheet('border: 0; background-color: #00f0f0ff;')
        
        self.inputEdit = QLineEdit(parent=self.PrintTemp)
        self.inputEdit.setGeometry(QtCore.QRect(510, 11, 180, 30))
        self.inputEdit.setObjectName("inputEdit")
        self.inputEdit.setPlaceholderText("Nhập...")
        self.inputEdit.textChanged.connect(lambda text: self.for_seclect_items(lambda i, _, code: [text, code]))
        
        self.label_mavach = QLabel(parent=self.PrintTemp)
        self.label_mavach.setGeometry(QtCore.QRect(790, 8, 100, 30))
        self.label_mavach.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_mavach.setObjectName("label_mavach")
        
        self.inputMavach = QSpinBox(parent=self.PrintTemp)
        self.inputMavach.setGeometry(QtCore.QRect(880, 11, 60, 30))
        self.inputMavach.setObjectName("inputCode")
        self.inputMavach.setPrefix("")
        self.inputMavach.setRange(0, 999)
        self.loadMavach()
        self.inputMavach.textChanged.connect(self.reLoadMavach)
        
        self.printButton = QPushButton(parent=self.PrintTemp)
        self.printButton.setGeometry(QtCore.QRect(*main_button))
        self.printButton.setObjectName("printButton")
        self.printButton.clicked.connect(self.print_selected_rows)
        self.printButton.setToolTip('Ctrl P')
        
        main_button[0] += main_button[2] + 5
        self.reprintButton = QPushButton(parent=self.PrintTemp)
        self.reprintButton.setGeometry(QtCore.QRect(*main_button))
        self.reprintButton.setObjectName("reprintButton")
        self.reprintButton.clicked.connect(self.print_history_selected_rows)
        self.reprintButton.setToolTip('Ctrl Shirt P')
        
        main_button[0] += main_button[2] + 5
        self.saveBtn = QPushButton(parent=self.PrintTemp)
        self.saveBtn.setGeometry(QtCore.QRect(*main_button))
        self.saveBtn.setObjectName("saveBtn")
        self.saveBtn.clicked.connect(self.save)
        self.saveBtn.setToolTip('Ctrl S')
        
        self.setting_button = QPushButton(parent=self.PrintTemp)
        self.setting_button.setGeometry(QtCore.QRect(1150, 480, 90, 31))
        self.setting_button.setObjectName("commandLinkButton")
        self.setting_button.clicked.connect(self.settingEvent)
        self.setting_button.setStyleSheet('border: 0; background-color: #00f0f0ff;')
        
        self.labelColumn = QLabel(parent=self.PrintTemp)
        self.labelColumn.setGeometry(QtCore.QRect(1085, 15, 120, 26))
        self.labelColumn.setObjectName("labelColumn")
        
        self.newColumn = QLineEdit(parent=self.PrintTemp)
        self.newColumn.setGeometry(QtCore.QRect(1170, 10, 200, 30))
        self.newColumn.setObjectName("newColumn")
        self.newColumn.setPlaceholderText("Tên cột mới...")
        self.newColumn.textChanged.connect(self.changeNewColumnName)
        
        self.web_view = QLabel(parent=self.PrintTemp)
        self.web_view.setGeometry(QtCore.QRect(1080, 50, 302, 422))
        self.web_view.setObjectName("web_view")
        self.web_view.setStyleSheet('border: 1px solid #bbb; background-color: #d6d6d6;')
        
        if (not self.rmb.get(base_name, 'none') == 'none'):
            height = self.rmb[base_name]['height']
            data = self.rmb[base_name]['data']
        else:
            height = 0
            data = []
        self.temp_view = QLabel(parent=self.PrintTemp)
        self.temp_view.setGeometry(QtCore.QRect(1081, 51 + (420 - height)//2, 300, height))
        self.temp_view.setObjectName("temp_view")
        self.temp_view.setStyleSheet('border: 2px solid gray; background-color: white;')
        self.design_temp(data)
        
        self.fixlist = QLabel(parent=self.PrintTemp)
        self.fixlist.setGeometry(QtCore.QRect(15, 480, 50, 25))
        self.fixlist.setObjectName("fixlist")
        self.fixlist.setText('<i style="font-size: 14px; color: blue;">Fix:</i>')
        
        self.fix_list = []
        self.fix_x = 0
        def add_fix(v):
            button = QPushButton(parent=self.PrintTemp)
            button.setGeometry(QtCore.QRect(40 + self.fix_x, 480, len(v) * 10 + 6, 25))
            button.setObjectName(v)
            button.setText(v)
            button.setStyleSheet('font-size: 14px; border: 0; background-color: #00000000; color: blue; font-style: italic; margin: 2px;')
            self.fix_list.append(button)
            self.fix_x += len(v) * 7 + 13
            
        for val in rmb['sizelist']:
            add_fix(val)
            
        self.fix_list[0].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][i % 19], code]))
        self.fix_list[1].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 1) % 19], code]))
        self.fix_list[2].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 2) % 19], code]))
        self.fix_list[3].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 3) % 19], code]))
        self.fix_list[4].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 4) % 19], code]))
        self.fix_list[5].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 5) % 19], code]))
        self.fix_list[6].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 6) % 19], code]))
        self.fix_list[7].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 7) % 19], code]))
        self.fix_list[8].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 8) % 19], code]))
        self.fix_list[9].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 9) % 19], code]))
        self.fix_list[10].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 10) % 19], code]))
        self.fix_list[11].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 11) % 19], code]))
        self.fix_list[12].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 12) % 19], code]))
        self.fix_list[13].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 13) % 19], code]))
        self.fix_list[14].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 14) % 19], code]))
        self.fix_list[15].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 15) % 19], code]))
        self.fix_list[16].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 16) % 19], code]))
        self.fix_list[17].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 17) % 19], code]))
        self.fix_list[18].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [rmb['sizelist'][(i + 18) % 19], code]))
        self.fix_list[19].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [f'SIZE {text}', code]))
        self.fix_list[20].clicked.connect(lambda: self.for_seclect_items(lambda i, text, code: [f'({text})', code]))
        
        #set Text
        self.PrintTemp.setWindowIcon(QIcon(iconPath))
        self.label.setText("Tìm Kiếm:")
        self.label_edit.setText("Chỉnh Sửa:")
        self.commandLinkButton.setText("DESIGN")
        self.label_mavach.setText("Mã Vạch:")
        self.printButton.setText("In Ra")
        self.reprintButton.setText("In Lại")
        self.labelColumn.setText("Tên cột mới:")
        self.saveBtn.setText("Lưu")
        self.setting_button.setText("Setting")
        
        QtCore.QMetaObject.connectSlotsByName(self.PrintTemp)
        self.load_excel_data(input_file)
        self.load_history()
        self.PrintTemp.setWindowTitle("Print Temp - " + file_name)
        self.tableWidget.cellClicked.connect(self.update_line_edit)
        self.tableWidget.cellChanged.connect(self.cell_change)
        self.PrintTemp.keyPressEvent = self.keyPressEvent
        
    def settingEvent(self):
        self.setting = SettingWindow()
        self.setting.colorPicked.connect(self.set_color)
        self.setting.show()
        
    def set_color(self, color):
        pass
        
    def on_selection_change(self, table, top=False): 
        selected_items_rows = [item.row() for item in table.selectedItems()]
        row_count = table.rowCount() 
        column_count = table.columnCount()
        if top and len(selected_items_rows) > 0:
            my_row = self.tableWidget.rawValue[selected_items_rows[-1]]
            headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if self.tableWidget.horizontalHeaderItem(col)]
            my_dict = {key: val for key, val in zip(headers, my_row)}
            self.design_temp(my_dict)
        for row in range(row_count): 
            for col in range(column_count): 
                item = table.item(row, col) 
                if item:
                    if row in selected_items_rows:
                        item.setBackground(QColor(0, 135, 255, 50))
                        # item.setForeground(QColor(0, 135, 255))
                    else:
                        item.setBackground(QColor(0, 0, 0, 0))
                        # item.setForeground(QColor(0, 0, 0))
                        
    def keyPressEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier and event.key() == Qt.Key.Key_P:
            self.print_history_selected_rows()
        elif event.key() == Qt.Key.Key_P:
            self.print_selected_rows()
        elif event.key() == Qt.Key.Key_S:
            self.save()
        elif event.key() == Qt.Key.Key_Z:
            self.tableWidget.undo_stacking = True
            self.tableWidget.undo(self.tableWidget.undo_stack, self.tableWidget.redo_stack)
        elif event.key() == Qt.Key.Key_Y:
            self.undo_stacking = True
            self.tableWidget.undo(self.tableWidget.redo_stack, self.tableWidget.undo_stack)
        elif event.modifiers() == Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_O:
            subprocess.run(f'bartend.exe "{btwFile}"', startupinfo=startupinfo, shell=True)
        
    def cell_change(self, row, col):
        self.PrintTemp.setWindowTitle("Print Temp - " + file_name + "*")
        if self.tableWidget.undo_stacking == False:
            if self.set_end_item:
                for item in self.set_undo:
                    newItem = self.tableWidget.rawValue[item['row']][item['col']]
                    if newItem != item['val']:
                        self.tableWidget.undo_stack.append({
                            'style': 'change_item',
                            'val': item['val'],
                            'row': item['row'],
                            'col': item['col'],
                            'n': len(self.set_undo)
                        })
        
    def del_row(self):
        self.tableWidget.remove_row()
        self.PrintTemp.setWindowTitle("Print Temp - " + file_name + "*")
            
    def del_col(self):
        self.tableWidget.remove_column()
        self.PrintTemp.setWindowTitle("Print Temp - " + file_name + "*")
        
    def load_excel_data(self, file_path: str):
        self.df = pd.read_excel(file_path).fillna('')
        self.tableWidget.setRowCount(self.df.shape[0])
        self.tableWidget.setColumnCount(self.df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(self.df.columns.tolist())
        self.tableWidget.rawValue = self.df.to_numpy().tolist()
        for i, row in self.df.iterrows():
            for j, value in enumerate(row):
                self.tableWidget.setItem(i, j, QTableWidgetItem(encode_value(value, rmb)))
                
    def load_history(self):
        headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if self.tableWidget.horizontalHeaderItem(col)]
        self.tableWidget_history.setRowCount(len(self.rmb[base_name]['history']))
        self.tableWidget_history.setColumnCount(len(headers) + 1)
        self.tableWidget_history.setHorizontalHeaderLabels([*headers, 'History'])
        for i, row in enumerate(self.rmb[base_name]['history'][::-1]):
            for j in range(self.tableWidget_history.columnCount()):
                value = row.get(self.tableWidget_history.horizontalHeaderItem(j).text(), "")
                self.tableWidget_history.setItem(i, j, QTableWidgetItem(value))
    
    def reload_history(self):
        old_len = len(self.rmb[base_name]['history'])
        headers_history = [self.tableWidget_history.horizontalHeaderItem(col).text() for col in range(self.tableWidget_history.columnCount()) if self.tableWidget_history.horizontalHeaderItem(col)]
        setList = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if not self.tableWidget.horizontalHeaderItem(col).text() in headers_history]
        if len(setList) > 0:
            column_count_history = self.tableWidget_history.columnCount()
            self.tableWidget_history.setColumnCount(column_count_history + len(setList))
            for i, nhead in enumerate(setList):
                self.tableWidget_history.setHorizontalHeaderItem(column_count_history - 1 + i, QTableWidgetItem(nhead))
            self.tableWidget_history.setHorizontalHeaderItem(column_count_history + len(setList) - 1, QTableWidgetItem('History'))
                
        if old_len > history_n:
            del self.rmb[base_name]['history'][0:(old_len - history_n)]
        
        new_len = len(self.rmb[base_name]['history'])
        oldrowlen = self.tableWidget_history.rowCount()
        
        if oldrowlen < new_len:
            for ind in range(new_len - oldrowlen):
                self.tableWidget_history.insertRow(oldrowlen + ind)
        
        set_checkColumn = [True for i in range(self.tableWidget_history.columnCount())]
        for i, row in enumerate(self.rmb[base_name]['history'][::-1]):
            for j in range(self.tableWidget_history.columnCount()):
                value = row.get(self.tableWidget_history.horizontalHeaderItem(j).text(), "")
                self.tableWidget_history.setItem(i, j, QTableWidgetItem(value))
                if (not value == "") and set_checkColumn[j] == True: set_checkColumn[j] = False
                
        setRemove = [i for i, check in enumerate(set_checkColumn) if check]
        for col in sorted(setRemove, reverse=True):
            self.tableWidget_history.removeColumn(col)

    def changeNewColumnName(self, text):
        self.tableWidget.newColumn_text = text
    
    def search(self, text: str):
        search_term = text.lower()
        for row in range(self.tableWidget.rowCount()):
            match = search_term in " | ".join([self.tableWidget.item(row, col).text().lower() for col in range(self.tableWidget.columnCount())])
            self.tableWidget.setRowHidden(row, not match)

            
    def print_selected_rows(self):
        selected_items = self.tableWidget.selectedItems()
        rrows = [item.row() for item in selected_items]
        rows = list(set(rrows))
        rows.sort()
        data = []
        now = datetime.datetime.now()
        time = now.strftime('%H:%M - %d/%m/%Y')
        headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if self.tableWidget.horizontalHeaderItem(col)]
        for row in rows:
            if not self.tableWidget.isRowHidden(row):
                row_data = [self.tableWidget.item(row, col).text() for col in range(self.tableWidget.columnCount())]
                data.append(row_data)
                self.rmb[base_name]['history'].append({key: val for key, val in zip([*headers, "History"], [*row_data, time])})
        
        if len(data) > 0:
            df_selected = pd.DataFrame(data, columns=headers)
            df_selected.to_excel(outputFile, index=False)
            subprocess.run(cmd, startupinfo=startupinfo, shell=True)
            self.reload_history()
            self.save()
            
    def print_history_selected_rows(self):
        selected_items = self.tableWidget_history.selectedItems()
        rrows = [item.row() for item in selected_items]
        rows = list(set(rrows))
        rows.sort()
        data = []
        now = datetime.datetime.now()
        time = now.strftime('%H:%M - %d/%m/%Y')
        headers = [self.tableWidget_history.horizontalHeaderItem(col).text() for col in range(self.tableWidget_history.columnCount()) if self.tableWidget_history.horizontalHeaderItem(col)]
        for row in rows:
            row_data = [self.tableWidget_history.item(row, col).text() for col in range(self.tableWidget_history.columnCount())]
            data.append(row_data)
            row_data[-1] = time
            self.rmb[base_name]['history'].append({key: val for key, val in zip([*headers, "History"], [*row_data, time])})
        
        if len(data) > 0:
            df_selected = pd.DataFrame(data, columns=headers)
            df_selected.to_excel(outputFile, index=False)
            subprocess.run(cmd, startupinfo=startupinfo, shell=True)
            self.reload_history()
            self.save()
            
    def update_line_edit(self, row: int, column: int):
        item_raw = str(self.tableWidget.rawValue[row][column])
        self.rechangeItem = False
        if len(self.tableWidget.selectedItems()) == 1:
            textlist = item_raw.split(" #")
            text = textlist[0]
            code = len(textlist) > 1 and textlist[1] or 0
            self.inputEdit.setText(text)
            self.inputCode.setValue(int(code))
            # else:
            #     self.inputEdit.setText(item_raw)
            #     self.inputCode.setValue(0)
        
    def loadMavach(self):
        if not self.rmb.get(base_name, 'none') == 'none':
            self.inputMavach.setValue(self.rmb[base_name]['mavach'])
            rmb["mavach"] = self.rmb[base_name]['mavach']
        
            
    def reLoadMavach(self):
        rmb["mavach"] = self.inputMavach.value()
        self.rmb[base_name]['mavach'] = self.inputMavach.value()
        selected_items = self.tableWidget.selectedItems()
        my_row = self.tableWidget.rawValue[len(selected_items) > 0 and selected_items[0].row() or 0]
        headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if self.tableWidget.horizontalHeaderItem(col)]
        my_dict = {key: val for key, val in zip(headers, my_row)}
        self.design_temp(my_dict)
        for row in range(self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                text = encode_value(self.tableWidget.rawValue[row][col], rmb)
                self.tableWidget.setItem(row, col, QTableWidgetItem(text))

    def design_temp(self, mydict: dict):
        if not self.rmb.get(base_name, 'none') == 'none':
            mydict['MaVach'] = 'mavach'
            mydict['Random'] = random.randint(0, 9)
            mydict['code1'] = "font-family: Arial; font-size: 11px; font-weight: 400;"
            mydict['code2'] = "font-family: Code 128; font-size: 45px; font-weight: 400;"
            mydict['code3'] = "font-family: c39hrp24dhtt; font-size: 50px; font-weight: 400;"
            
            for key, val in self.rmb[base_name]['data'].items():
                mydict[key] = mydict.get(key, f'{{ {key} }}')
                
            newdict = {key: encode_value(value, rmb) for key, value in mydict.items()}
            if base_name == "SP_HANG_HOA.xlsx" or base_name == "SP_HANG_HOA - Copy.xlsx":
                text = newdict['SIZE'] + newdict['TYPE']
                if len(text) > 14:
                    newdict['sizeStyle'] = "margin: 20px 0; font-size: 19px; word-spacing: 0px;"
                elif newdict['SIZE'] == 'QT':
                    newdict['sizeStyle'] = "margin: 20px 0; font-size: 28px; word-spacing: 0px;"
                elif len(text) > 7:
                    newdict['sizeStyle'] = "margin: 20px 0; font-size: 26px; word-spacing: 0;"
                    newdict['SIZE'] = ' '.join(newdict['SIZE'])
                    newdict['TYPE'] = ' '.join(newdict['TYPE'])
                else:
                    newdict['sizeStyle'] = "margin: 20px 0; font-size: 29px; word-spacing: 8px;"
                    newdict['SIZE'] = ' '.join(newdict['SIZE'])
                    newdict['TYPE'] = ' '.join(newdict['TYPE'])
            elif base_name == "MaVach.xlsx":
                newdict['sizeStyle'] = "font-family: VNF-ITC Lubalin Graph; font-size: 32px; word-spacing: 15px; font-weight: bold; width: 70%;"
            self.temp_view.setText(self.rmb[base_name]['design'] % newdict)
            self.rmb[base_name]['data'] = mydict
            
    def for_seclect_items(self, fn):
        selected_items = self.tableWidget.selectedItems()
        i = 0
        self.set_undo = []
        for item in selected_items:
            if not self.tableWidget.isRowHidden(item.row()):
                rawStr = str(self.tableWidget.rawValue[item.row()][item.column()])
                textlist = rawStr.split(" #")
                if i == len(selected_items)-1:
                    self.set_end_item = True
                self.set_undo.append({
                    'row': item.row(),
                    'col': item.column(),
                    'val': rawStr
                })
                code = len(textlist) > 1 and " #" + textlist[1] or ""
                text = str(textlist[0])
                rmb['i'] = i
                rmb['code'] = self.inputCode.value()
                out = fn(i, text, code)
                if out[1] == " #0": out[1] = ""
                self.tableWidget.rawValue[item.row()][item.column()] = ''.join(out)
                out[0] = encode_value(out[0], rmb)
                item.setText(''.join(out))
                i += 1
        headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if self.tableWidget.horizontalHeaderItem(col)]
        my_dict = {key: val for key, val in zip(headers, self.tableWidget.rawValue[selected_items[0].row()])}
        self.design_temp(my_dict)
        self.set_end_item = False
        self.set_undo = []
          
    def save(self):
        headers = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount()) if self.tableWidget.horizontalHeaderItem(col)]
        df = pd.DataFrame(self.tableWidget.rawValue, columns=headers)
        df.to_excel(input_file, index=False)
        self.PrintTemp.setWindowTitle("Print Temp - " + file_name)
        save_json(jsPath, self.rmb)
            
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    win = QtWidgets.QMainWindow()
    ui = Ui_PrintTemp()
    ui.setupUi(win)
    win.show()
    sys.exit(app.exec())