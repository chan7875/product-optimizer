import sys
import os
import subprocess
import pandas as pd
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTableView, QFileDialog, QTabWidget, QLabel, 
                             QLineEdit, QMessageBox, QHeaderView, QAbstractItemView,
                             QInputDialog, QDialog, QTextEdit, QTableWidget, QTableWidgetItem,
                             QDateEdit, QSplitter, QTreeWidget, QTreeWidgetItem)
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QDate
from PyQt6.QtGui import QColor, QFont
import shutil
import csv

class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super(PandasModel, self).__init__()
        self._data = data
        self._date_columns = []
        self._weekend_columns = []
        self._identify_date_columns()

    def _identify_date_columns(self):
        self._date_columns = []
        self._weekend_columns = []
        for col_idx, col_name in enumerate(self._data.columns):
            try:
                s_col = str(col_name).split(' ')[0]
                dt = pd.to_datetime(s_col, errors='coerce')
                if not pd.isna(dt):
                    self._date_columns.append(col_idx)
                    if dt.dayofweek >= 5:
                        self._weekend_columns.append(col_idx)
            except:
                pass

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
            val = self._data.iloc[row, col]
            if pd.isna(val):
                return ""
            return str(val)

        if role == Qt.ItemDataRole.BackgroundRole:
            if col in self._weekend_columns:
                return QColor(220, 220, 220)
            if col in self._date_columns:
                val = self._data.iloc[row, col]
                try:
                    f_val = float(val)
                    if f_val > 0:
                        return QColor(255, 255, 200)
                except:
                    pass
        return None

    def setData(self, index, value, role):
        if role == Qt.ItemDataRole.EditRole:
            row = index.row()
            col = index.column()
            try:
                if value == "":
                    self._data.iloc[row, col] = None
                else:
                    try:
                        f_val = float(value)
                        if f_val.is_integer():
                            self._data.iloc[row, col] = int(f_val)
                        else:
                            self._data.iloc[row, col] = f_val
                    except ValueError:
                        self._data.iloc[row, col] = value
                
                self.dataChanged.emit(index, index, [Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole])
                return True
            except Exception as e:
                print(f"SetData Error: {e}")
                return False
        return False

    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Orientation.Vertical:
                return str(self._data.index[section] + 1)
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

    def insertRows(self, position, rows, parent=QModelIndex()):
        self.beginInsertRows(parent, position, position + rows - 1)
        empty_row = pd.Series([None]*self.columnCount(), index=self._data.columns)
        new_df = pd.DataFrame([empty_row])
        self._data = pd.concat([self._data.iloc[:position], new_df, self._data.iloc[position:]]).reset_index(drop=True)
        self.endInsertRows()
        return True
    
    def get_dataframe(self):
        return self._data
    
    def set_dataframe(self, df):
        self.beginResetModel()
        self._data = df
        self._identify_date_columns()
        self.endResetModel()


# Import logic from existing scripts
import calculate_schedule


class SMDVerificationTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        
        # Top Bar Container
        self.top_container = QWidget()
        self.top_container.setFixedHeight(50)
        top_bar = QHBoxLayout(self.top_container)
        top_bar.setContentsMargins(10, 5, 10, 5) 
        
        self.btn_load_folder = QPushButton("Load Folder")
        self.btn_load_folder.clicked.connect(self.load_folder)
        
        self.btn_run_smd = QPushButton("SMD Pro 실행")
        self.btn_run_smd.clicked.connect(self.run_smd_pro)
        
        self.lbl_path = QLabel("No folder selected")
        
        top_bar.addWidget(self.btn_load_folder)
        top_bar.addWidget(self.btn_run_smd)
        top_bar.addWidget(self.lbl_path)
        top_bar.addStretch()
        
        self.layout.addWidget(self.top_container)
        
        # Splitter (Left: Tree, Right: Table)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: Tree Widget (Filter/Navigation)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("SMD Code / PCB Code")
        self.tree.setFixedWidth(250)
        self.splitter.addWidget(self.tree)
        
        # Right: Table Widget (Details)
        self.table = QTableWidget()
        cols = ["Check", "PCB Code", "Rev", "SMD Code", "Model Name", "PCB Size", 
                "Neutral File", "BOM File", "Gerber File", "Matr Count", "Path"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.splitter.addWidget(self.table)
        
        # Set Splitter Stretch
        self.splitter.setStretchFactor(1, 1)
        
        self.layout.addWidget(self.splitter)
        self.setLayout(self.layout)
        
        self.json_data_list = []
        
        self.tree.itemClicked.connect(self.filter_table)
        self.tree.header().setSectionsClickable(True)
        self.tree.header().sectionClicked.connect(self.reset_filter)

    def load_folder(self):
        default_dir = r"Y:\CadDesign\Manufacture\NW\Design_25"
        if not os.path.exists(default_dir):
            default_dir = ""
            
        folder_path = QFileDialog.getExistingDirectory(self, "Select Root Directory", default_dir)
        if folder_path:
            self.lbl_path.setText(folder_path)
            self.scan_directory(folder_path)

    def scan_directory(self, root_path):
        self.table.setRowCount(0)
        self.tree.clear()
        self.json_data_list = []
        
        # Recursive walk
        for root, dirs, files in os.walk(root_path):
            if "jsonInfo.txt" in files:
                full_path = os.path.join(root, "jsonInfo.txt")
                data = self.parse_json_info(full_path)
                if data:
                    data['FolderPath'] = root
                    self.json_data_list.append(data)
        
        self.populate_ui()
        QMessageBox.information(self, "Done", f"Found {len(self.json_data_list)} jsonInfo.txt files.")

    def parse_json_info(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = json.load(f)
                
            basic = content.get('basicInfo', {})
            cad = content.get('cadFileInfo', {})
            
            # Extract Matr List (BOMs) - Check multiple locations
            matr_list = cad.get('matrList')
            if not matr_list:
                matr_list = content.get('matrList')
            if not matr_list:
                matr_list = basic.get('matrList')
            
            if matr_list is None: 
                matr_list = []
                
            bom_files = [m.get('matrFileNm', '') for m in matr_list if isinstance(m, dict)]
            bom_str = "\n".join(bom_files)
            
            return {
                'pcbCode': basic.get('pcbCode', ''),
                'rev': basic.get('seq', ''),
                'smdCode': basic.get('smdCode', ''),
                'smdNm': basic.get('smdNm', ''),
                'pcbSize': basic.get('pcbSize', ''),
                'neutralFileNm': cad.get('neutralFileNm', ''),
                'gerberFileNm': cad.get('gerberFileNm', ''),
                'bomFiles': bom_str,
                'matrCount': len(matr_list)
            }
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return None

    def reset_filter(self, index):
        # Show all rows
        self.tree.clearSelection()
        for row in range(self.table.rowCount()):
            self.table.setRowHidden(row, False)

    def filter_table(self, item, column):
        data = item.data(0, Qt.ItemDataRole.UserRole)
        if not data:
            # Show all if no data (optional, or do nothing)
            for row in range(self.table.rowCount()):
                self.table.setRowHidden(row, False)
            return
            
        filter_type, filter_value = data
        
        for row in range(self.table.rowCount()):
            should_show = False
            if filter_type == "SMD":
                # Col 3 is SMD Code
                table_item = self.table.item(row, 3)
                if table_item and table_item.text() == filter_value:
                    should_show = True
            elif filter_type == "PCB":
                # Col 1 is PCB Code
                table_item = self.table.item(row, 1)
                # Note: Table might have duplicates for PCB Code if multiple SMDs? 
                # Actually specific row logic:
                # If filter is PCB, we match PCB Code.
                if table_item and table_item.text() == filter_value:
                    should_show = True
            
            self.table.setRowHidden(row, not should_show)

    def run_smd_pro(self):
        checked_rows = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.checkState() == Qt.CheckState.Checked:
                # Get Path from column 10
                path_item = self.table.item(row, 10)
                if path_item:
                    checked_rows.append(path_item.text())
        
        if not checked_rows:
            QMessageBox.warning(self, "Warning", "선택된 항목이 없습니다.")
            return

        # Executable Path
        exe_path = r"C:\Program Files (x86)\Pentacube\Cubic\Manufacture\CubicSMT\NW\CubicSMT.exe"
        
        if not os.path.exists(exe_path):
             # For safety/testing, warn if not found, but try to run strict logic as requested or just warn?
             # User environment likely has it. I'll add a check but maybe just log it if missing?
             # Actually user said "Code execute this". 
             # I will warn if missing, but proceed or return?
             # Better to warn and return to avoid crash, or just print.
             QMessageBox.warning(self, "Error", f"Executable not found:\n{exe_path}")
             return

        executed_count = 0
        for path in checked_rows:
            json_path = os.path.join(path, "jsonInfo.txt")
            
            # Command: [exe, -nwJsoninfo, json_path, -nwMounter]
            cmd = [exe_path, "-nwJsoninfo", json_path, "-nwMounter"]
            
            try:
                subprocess.Popen(cmd)
                executed_count += 1
            except Exception as e:
                print(f"Error executing command for {path}: {e}")
                
        if executed_count > 0:
            QMessageBox.information(self, "Success", f"Executed CubicSMT for {executed_count} items.")

    def populate_ui(self):
        # Populate Table
        self.table.setRowCount(len(self.json_data_list))
        
        # Tree Structure: Map by SMD Code? Or PCB Code? 
        # User Image shows "SMD Code Criteria"
        tree_map = {} 
        
        for i, data in enumerate(self.json_data_list):
            # Table Row
            # Checkbox item for first column?
            item_check = QTableWidgetItem("")
            item_check.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            item_check.setCheckState(Qt.CheckState.Unchecked)
            self.table.setItem(i, 0, item_check)
            
            self.table.setItem(i, 1, QTableWidgetItem(data['pcbCode']))
            self.table.setItem(i, 2, QTableWidgetItem(data['rev']))
            self.table.setItem(i, 3, QTableWidgetItem(data['smdCode']))
            self.table.setItem(i, 4, QTableWidgetItem(data['smdNm']))
            self.table.setItem(i, 5, QTableWidgetItem(data['pcbSize']))
            self.table.setItem(i, 6, QTableWidgetItem(data['neutralFileNm']))
            self.table.setItem(i, 7, QTableWidgetItem(data['bomFiles']))
            self.table.setItem(i, 8, QTableWidgetItem(data['gerberFileNm']))
            self.table.setItem(i, 9, QTableWidgetItem(str(data['matrCount'])))
            self.table.setItem(i, 10, QTableWidgetItem(data['FolderPath']))

            # Tree Population
            key = data['smdCode'] 
            if not key: key = "Unknown"
            if key not in tree_map:
                tree_map[key] = []
            tree_map[key].append(data)
            
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        
        # Build Tree
        self.tree.clear()
        for smd_code, items in tree_map.items():
            parent = QTreeWidgetItem(self.tree)
            parent.setText(0, f"{smd_code} ({len(items)})")
            parent.setData(0, Qt.ItemDataRole.UserRole, ("SMD", smd_code))
            
            # Use a set to avoid duplicate PCB codes under same SMD if any
            # But usually one row per file/item.
            
            for item in items:
                child = QTreeWidgetItem(parent)
                child.setText(0, item['pcbCode'])
                child.setData(0, Qt.ItemDataRole.UserRole, ("PCB", item['pcbCode']))

class ScheduleTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()
        
        # Buttons Setup
        btn_layout = QHBoxLayout()
        self.btn_load = QPushButton("Load Excel Schedule")
        self.btn_load.clicked.connect(self.load_excel)
        self.btn_save = QPushButton("Save/Export Excel")
        self.btn_save.clicked.connect(self.save_excel)
        self.btn_add_row = QPushButton("Add Row")
        self.btn_add_row.clicked.connect(self.add_row)
        self.btn_calc = QPushButton("Calculate Production Time")
        self.btn_calc.clicked.connect(self.calculate_time)
        
        btn_layout.addWidget(self.btn_load)
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_add_row)
        btn_layout.addWidget(self.btn_calc)
        
        # New Button for showing changes
        self.btn_show_changes = QPushButton("Show Last Changes")
        self.btn_show_changes.setEnabled(False)
        self.btn_show_changes.clicked.connect(self.show_changes_dialog)
        btn_layout.addWidget(self.btn_show_changes)

        btn_layout.addStretch()
        
        self.layout.addLayout(btn_layout)
        
        # Setup Times Input Group
        setup_group = QHBoxLayout()
        setup_group.addWidget(QLabel("<b>Setup Times (min):</b>"))
        
        self.setup_inputs = {}
        defaults = {'S01': '40', 'S02': '13', 'S03': '40', 'S04': '13'}
        
        for line in ['S01', 'S02', 'S03', 'S04']:
            setup_group.addWidget(QLabel(f"{line}:"))
            le = QLineEdit(defaults[line])
            le.setFixedWidth(40)
            setup_group.addWidget(le)
            self.setup_inputs[line] = le
            
        setup_group.addStretch()
        self.layout.addLayout(setup_group)
        
        # Quad-View Setup (Frozen Rows & Cols)
        # Structure:
        # TL (Fixed) | TR (H-Scroll)
        # BL (V-Scroll)| BR (HV-Scroll)
        
        self.model_summary = PandasModel(pd.DataFrame())
        self.model_main = PandasModel(pd.DataFrame())
        
        # Grid Layout for 4 tables
        from PyQt6.QtWidgets import QGridLayout
        grid = QGridLayout()
        grid.setSpacing(0)
        grid.setContentsMargins(0,0,0,0)
        
        # 1. Top Left (Fixed)
        self.table_tl = QTableView()
        self.table_tl.setModel(self.model_summary)
        self.table_tl.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table_tl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table_tl.setAlternatingRowColors(True)
        self.table_tl.setStyleSheet("border-right: 1px solid gray; border-bottom: 1px solid gray;")
        
        # 2. Top Right (H-Scroll, V-Fixed)
        self.table_tr = QTableView()
        self.table_tr.setModel(self.model_summary)
        self.table_tr.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table_tr.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff) # Controlled by BR
        self.table_tr.setAlternatingRowColors(True)
        self.table_tr.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table_tr.setStyleSheet("border-bottom: 1px solid gray;")
        
        # 3. Bottom Left (V-Scroll, H-Fixed)
        self.table_bl = QTableView()
        self.table_bl.setModel(self.model_main)
        self.table_bl.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff) # Controlled by BR
        self.table_bl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table_bl.setAlternatingRowColors(True)
        self.table_bl.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel) # Ideally fixed
        self.table_bl.setStyleSheet("border-right: 1px solid gray;")
        
        # 4. Bottom Right (Main Scrollable)
        self.table_br = QTableView()
        self.table_br.setModel(self.model_main)
        self.table_br.setAlternatingRowColors(True)
        self.table_br.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        
        # Sync Scrollbars
        # Vertical: BL <-> BR
        self.table_bl.verticalScrollBar().valueChanged.connect(self.table_br.verticalScrollBar().setValue)
        self.table_br.verticalScrollBar().valueChanged.connect(self.table_bl.verticalScrollBar().setValue)
        
        # Horizontal: TR <-> BR
        self.table_tr.horizontalScrollBar().valueChanged.connect(self.table_br.horizontalScrollBar().setValue)
        self.table_br.horizontalScrollBar().valueChanged.connect(self.table_tr.horizontalScrollBar().setValue)
        
        # Sync Column Resizing (Crucial for alignment)
        # Connect horizontal headers of TR and BR
        self.table_tr.horizontalHeader().sectionResized.connect(self.sync_resize_tr_br)
        self.table_br.horizontalHeader().sectionResized.connect(self.sync_resize_br_tr)
        
        # Connect TL and BL
        self.table_tl.horizontalHeader().sectionResized.connect(self.sync_resize_tl_bl)
        self.table_bl.horizontalHeader().sectionResized.connect(self.sync_resize_bl_tl)
        
        # Connect Double Click for Detail View
        self.table_tr.doubleClicked.connect(self.on_summary_double_click)

        # Add to Grid
        # Row 0: Summary
        # Row 1: Main
        # Col 0: Frozen Col
        # Col 1: Main Col
        
        grid.addWidget(self.table_tl, 0, 0)
        grid.addWidget(self.table_tr, 0, 1)
        grid.addWidget(self.table_bl, 1, 0)
        grid.addWidget(self.table_br, 1, 1)
        
        # Stretch factors
        grid.setColumnStretch(0, 1) # Frozen
        grid.setColumnStretch(1, 3) # Main
        grid.setRowStretch(0, 0)    # Summary (Fixed height)
        grid.setRowStretch(1, 1)    # Main (Expand)
        
        self.layout.addLayout(grid)
        self.setLayout(self.layout)
        self.current_filepath = ""

    # resizing slots
    def sync_resize_tr_br(self, index, old, new):
        self.table_br.setColumnWidth(index, new)
    def sync_resize_br_tr(self, index, old, new):
        self.table_tr.setColumnWidth(index, new)
    def sync_resize_tl_bl(self, index, old, new):
        self.table_bl.setColumnWidth(index, new)
    def sync_resize_bl_tl(self, index, old, new):
        self.table_tl.setColumnWidth(index, new)

    def on_summary_double_click(self, index):
        if not self.current_filepath: return
        
        # Identify Line: Row 0,1 -> S01; 2,3 -> S02...
        lines = ['S01', 'S02', 'S03', 'S04']
        line_idx = index.row() // 2
        if line_idx >= len(lines): return
        
        line_name = lines[line_idx]
        
        # Identify Date Column
        col_idx = index.column()
        if col_idx >= self.model_main.columnCount(): return
        
        # Check if actual date column
        # Using header text or trying to parse
        df = self.model_main.get_dataframe()
        col_name = df.columns[col_idx]
        try:
             s_col = str(col_name).split(' ')[0]
             # Expecting YYYY-MM-DD from our format logic
             dt = pd.to_datetime(s_col, errors='coerce')
             if pd.isna(dt): return
             date_str = dt.strftime('%Y-%m-%d')
        except: return
        
        self.show_detail_popup(line_name, col_idx, date_str, df)

    def show_detail_popup(self, line_name, date_col_idx, date_str, df):
        # 1. Identify Line Column Index
        line_col_idx = getattr(self, 'line_col_idx', None)
        if line_col_idx is None:
            for i, col in enumerate(df.columns):
                c_str = str(col).lower()
                if "line" in c_str or "생산라인" in c_str:
                    line_col_idx = i
                    break
        
        # 2. Setup Map
        setup_map = {}
        for ln, le in self.setup_inputs.items():
            try: setup_map[ln] = float(le.text())
            except: setup_map[ln] = 0.0
            
        details = []
        
        # 3. Iterate and Filter
        for r_idx, row in df.iterrows():
            qty_val = row.iloc[date_col_idx]
            if pd.isna(qty_val) or qty_val <= 0: continue
            
            # Check Line Match
            row_line = "Unknown"
            if line_col_idx is not None:
                val = str(row.iloc[line_col_idx]).strip()
                if val and val != 'nan': row_line = val
            
            if row_line != line_name: continue
            
            # Calculate
            res = calculate_schedule.calculate_time_for_row(row, qty_val, setup_map, line_col_idx)
            details.extend(res)
        
        # Calculate Statistics
        total_time = sum(d.get('Prod_Time', 0) for d in details)
        count = len(details)
        op_rate = (total_time / 480.0) * 100 if total_time > 0 else 0.0

        # 4. Show Dialog
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Production Details - {line_name} ({date_str})")
        dlg.resize(650, 500)
        lay = QVBoxLayout()
        
        header_lbl = QLabel(f"Line: {line_name} | Date: {date_str}")
        header_lbl.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        lay.addWidget(header_lbl)
        
        # INSERT SUMMARY STATS HERE
        stats_text = (f"Total Production Count (Items): {count}\n"
                      f"Total Production Time: {total_time:.0f} minutes\n"
                      f"Operation Rate (vs 480min): {op_rate:.1f}%")
        stats_lbl = QLabel(stats_text)
        stats_lbl.setFont(QFont("Consolas", 10))
        # Optional styling
        stats_lbl.setStyleSheet("background-color: #f9f9f9; border: 1px solid #ddd; padding: 5px;")
        lay.addWidget(stats_lbl)
        
        table = QTableWidget()
        cols = ["Item Code", "Layer", "Qty", "T/T", "Time (min)"]
        table.setColumnCount(len(cols))
        table.setHorizontalHeaderLabels(cols)
        table.setRowCount(len(details))
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        
        for i, item in enumerate(details):
            # Item Code
            table.setItem(i, 0, QTableWidgetItem(str(item.get('Item_Code', ''))))
            # Layer
            table.setItem(i, 1, QTableWidgetItem(str(item.get('Layer', ''))))
            # Qty
            table.setItem(i, 2, QTableWidgetItem(str(item.get('Qty', 0))))
            # T/T
            table.setItem(i, 3, QTableWidgetItem(str(item.get('Cycle_Time', 0))))
            # Time
            p_time = item.get('Prod_Time', 0)
            table.setItem(i, 4, QTableWidgetItem(f"{p_time:.2f}"))
            
        table.resizeColumnsToContents()
        lay.addWidget(table)
        
        # Removed bottom label as requested
        
        dlg.setLayout(lay)
        dlg.exec()

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Schedule Excel", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            try:
                # 1. Load Main Data
                df_test = pd.read_excel(file_path, header=None, nrows=40)
                header_row = 0
                if df_test.shape[0] > 36:
                     row_vals = [str(x) for x in df_test.iloc[36].values]
                     if any("Item" in x for x in row_vals) or any("Code" in x for x in row_vals):
                         header_row = 36
                
                self.df = pd.read_excel(file_path, header=header_row)
                
                # Format Date Columns to YYYY-MM-DD
                new_cols = []
                for col in self.df.columns:
                    try:
                        if isinstance(col, pd.Timestamp):
                             new_cols.append(col.strftime('%Y-%m-%d'))
                        else:
                            # Try parsing string
                            s_col = str(col)
                            dt = pd.to_datetime(s_col, errors='coerce')
                            # Ensure it's a valid recent date to avoid false positives (e.g. pure numbers)
                            if not pd.isna(dt) and dt.year > 2000:
                                new_cols.append(dt.strftime('%Y-%m-%d'))
                            else:
                                new_cols.append(col)
                    except:
                        new_cols.append(col)
                self.df.columns = new_cols

                self.model_main.set_dataframe(self.df)
                
                # Check for existing data comparison
                if hasattr(self, 'last_df') and self.last_df is not None:
                    # Perform comparison
                    diff_report = self.compare_schedules(self.last_df, self.df)
                    self.change_log = diff_report
                    
                    if diff_report != "No changes detected.":
                        self.btn_show_changes.setEnabled(True)
                        QMessageBox.information(self, "Update", "Schedule Updated! changes detected.\nClick 'Show Last Changes' to view details.")
                    else:
                         self.btn_show_changes.setEnabled(False)
                else:
                    self.btn_show_changes.setEnabled(False)

                # Store current as last for next time
                self.last_df = self.df.copy()
                
                self.current_filepath = file_path
                
                # 2. Determine Split Col (Array) - Moved up
                self.split_col = -1
                for i, col in enumerate(self.df.columns):
                    if "Array" in str(col) or "연배열" in str(col):
                        self.split_col = i
                        break
                if self.split_col == -1: self.split_col = 8
                if self.split_col >= len(self.df.columns): self.split_col = 2
                
                # 3. Calculate Summary Data (Pass split_col for label placement)
                self.calc_summary(self.df, self.split_col)
                
                # 4. Hide Headers for Frozen Cols in Summary Model
                # We do this by renaming columns in the summary dataframe to empty strings
                # Only for 0 to split_col
                summ_df = self.model_summary.get_dataframe()
                new_cols = list(summ_df.columns)
                for i in range(self.split_col + 1):
                    if i == self.split_col:
                        new_cols[i] = "생산라인"
                    else:
                        new_cols[i] = ""
                summ_df.columns = new_cols
                self.model_summary.set_dataframe(summ_df)
                
                # 5. Apply Frozen Column Logic (to all 4 tables)
                # Show all first
                for t in [self.table_tl, self.table_tr, self.table_bl, self.table_br]:
                    for c in range(t.model().columnCount()):
                        t.setColumnHidden(c, False)
                
                # Hide Right cols in Left Tables (Frozen)
                for c in range(self.split_col + 1, self.model_main.columnCount()):
                    self.table_tl.setColumnHidden(c, True)
                    self.table_bl.setColumnHidden(c, True)
                    
                # Hide Left cols in Right Tables (Scrollable)
                for c in range(0, self.split_col + 1):
                    self.table_tr.setColumnHidden(c, True)
                    self.table_br.setColumnHidden(c, True)
                
                # Adjust Widths
                # Resize Summary First to fit labels
                self.table_tl.resizeColumnsToContents()
                self.table_bl.resizeColumnsToContents()
                
                # Sync fixed widths to Top
                for c in range(0, self.split_col + 1):
                    # Use max width of TL or BL? TL has labels now.
                    w_tl = self.table_tl.columnWidth(c)
                    w_bl = self.table_bl.columnWidth(c)
                    w = max(w_tl, w_bl)
                    
                    # Add padding ONLY to Array column (split_col) where labels are
                    if c == self.split_col:
                        w += 40
                        
                    self.table_bl.setColumnWidth(c, w)
                    self.table_tl.setColumnWidth(c, w)
                    
                # Calculate Fixed Width for container (optional)
                w = self.table_bl.verticalHeader().width() + 2
                for c in range(0, self.split_col + 1):
                    w += self.table_bl.columnWidth(c)
                if w > 800: w = 800
                
                self.table_tl.setFixedWidth(w + 5)
                self.table_bl.setFixedWidth(w + 5)
                
                # Adjust Height of Summary Tables
                h = self.table_tl.verticalHeader().length() + self.table_tl.horizontalHeader().height() + 2
                self.table_tl.setFixedHeight(h)
                self.table_tr.setFixedHeight(h)
                
                QMessageBox.information(self, "Success", f"Loaded data from {file_path}")
            except Exception as e:
                import traceback
                traceback.print_exc()
                QMessageBox.critical(self, "Error", str(e))

    def compare_schedules(self, old_df, new_df):
        """
        Compare two dataframes based on Item Code (Col 0).
        Returns a formatted string report.
        """
        report = []
        
        # Helper to get Item Code map: Code -> Series (Row)
        def get_item_map(df):
            mapping = {}
            # Assume Item Code is always Col 0 as per logic
            for idx, row in df.iterrows():
                code = str(row.iloc[0]).strip()
                if code and code != 'nan':
                    # If duplicate codes exist, we take the first one or manage duplicates?
                    # For simplicty, assuming unique or last one overrides
                    mapping[code] = row
            return mapping

        old_map = get_item_map(old_df)
        new_map = get_item_map(new_df)
        
        old_keys = set(old_map.keys())
        new_keys = set(new_map.keys())
        
        # 1. Added
        added = new_keys - old_keys
        if added:
            report.append("=== [Added Items] ===")
            for code in sorted(added):
                report.append(f"{code} 추가")
            report.append("")
            
        # 2. Deleted
        deleted = old_keys - new_keys
        if deleted:
            report.append("=== [Deleted Items] ===")
            for code in sorted(deleted):
                report.append(f"{code} 삭제")
            report.append("")
            
        # 3. Modified (Quantity Changes)
        # We need to compare date columns quantity.
        # Identify Date Columns in New DF (assuming structure is similar)
        date_cols = []
        for col_idx, col_name in enumerate(new_df.columns):
            try:
                s_col = str(col_name).split(' ')[0]
                pd.to_datetime(s_col, errors='coerce') # Validation
                # We trust the column index alignment if templates are same
                date_cols.append(col_idx)
            except: pass
            
        common = old_keys & new_keys
        modified_list = []
        
        for code in common:
            row_old = old_map[code]
            row_new = new_map[code]
            
            changes = []
            
            # Compare each date column value
            for col_idx in date_cols:
                # Safety check if column exists in both
                if col_idx < len(row_old) and col_idx < len(row_new):
                    val_old = row_old.iloc[col_idx]
                    val_new = row_new.iloc[col_idx]
                    
                    # Normalize
                    try: v_o = float(val_old) 
                    except: v_o = 0.0
                    if pd.isna(v_o): v_o = 0.0
                    
                    try: v_n = float(val_new) 
                    except: v_n = 0.0
                    if pd.isna(v_n): v_n = 0.0
                    
                    if abs(v_o - v_n) > 0.001: # Update found
                        col_header = str(new_df.columns[col_idx]).split(' ')[0]
                        changes.append(f"{col_header}: {int(v_o)} -> {int(v_n)}")
            
            if changes:
                modified_list.append(f"[{code}] " + ", ".join(changes))
                
        if modified_list:
            report.append("=== [Modified Items (Qty)] ===")
            for item in modified_list:
                report.append(item)
            report.append("")
            
        if not report:
            return "No changes detected."
        
        return "\n".join(report)

    def get_production_data(self, date_str):
        """
        Retrieves production data for a specific date from the currently loaded dataframe.
        Returns a list of dictionaries (compatible with item_list.txt format).
        """
        if self.model_main.rowCount() == 0:
            return None, "No schedule loaded."
            
        df = self.model_main.get_dataframe()
        
        # Find Date Column
        date_col_idx = -1
        for i, col in enumerate(df.columns):
            s_col = str(col).split(' ')[0]
            try:
                dt = pd.to_datetime(s_col, errors='coerce')
                if not pd.isna(dt) and dt.strftime('%Y-%m-%d') == date_str:
                    date_col_idx = i
                    break
            except: pass
            
        if date_col_idx == -1:
            return None, f"Date '{date_str}' not found in schedule."

        # Setup Map
        setup_map = {}
        for ln, le in self.setup_inputs.items():
            try: setup_map[ln] = float(le.text())
            except: setup_map[ln] = 0.0
            
        # Line Column
        line_col_idx = getattr(self, 'line_col_idx', None)
        if line_col_idx is None:
            for i, col in enumerate(df.columns):
                c_str = str(col).lower()
                if "line" in c_str or "생산라인" in c_str:
                    line_col_idx = i
                    break

        all_items = []
        for r_idx, row in df.iterrows():
            qty_val = row.iloc[date_col_idx]
            if pd.isna(qty_val) or qty_val <= 0: continue
            
            res = calculate_schedule.calculate_time_for_row(row, qty_val, setup_map, line_col_idx)
            all_items.extend(res)
            
        return all_items, None



    def show_changes_dialog(self):
        if not hasattr(self, 'change_log') or not self.change_log:
            QMessageBox.information(self, "Info", "No change history available.")
            return
            
        dlg = QDialog(self)
        dlg.setWindowTitle("Schedule Changes")
        dlg.resize(500, 600)
        lay = QVBoxLayout()
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setText(self.change_log)
        text_edit.setFont(QFont("Consolas", 10))
        lay.addWidget(text_edit)
        
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        lay.addWidget(btn_close)
        
        dlg.setLayout(lay)
        dlg.exec()

    def calc_summary(self, df, label_col_idx):
        # Identify columns
        # Setup map
        setup_map = {}
        for line, le in self.setup_inputs.items():
            try:
                setup_map[line] = float(le.text())
            except:
                setup_map[line] = 0.0

        line_col_idx = None
        for i, col in enumerate(df.columns):
            c_str = str(col).lower()
            if "line" in c_str or "생산라인" in c_str:
                line_col_idx = i
                break
        
        # Summary Rows
        lines = ['S01', 'S02', 'S03', 'S04'] 
        
        # Date Cols
        date_cols = []
        for col_idx, col_name in enumerate(df.columns):
            try:
                s_col = str(col_name).split(' ')[0]
                dt = pd.to_datetime(s_col, errors='coerce')
                if not pd.isna(dt):
                    date_cols.append(col_idx)
            except:
                pass
        
        # Cache calculations
        col_line_totals = {} 
        
        for r_idx, row in df.iterrows():
             for d_col in date_cols:
                qty_val = row.iloc[d_col]
                if pd.isna(qty_val) or qty_val == 0:
                    continue
                
                res = calculate_schedule.calculate_time_for_row(row, qty_val, setup_map, line_col_idx)
                for item in res:
                    ln = item['Line']
                    t = item['Prod_Time']
                    key = (d_col, ln)
                    col_line_totals[key] = col_line_totals.get(key, 0) + t

        # Build Summary DF
        summ_rows = []
        num_cols = len(df.columns)
        
        for line in lines:
            row_time = [None] * num_cols
            row_util = [None] * num_cols
            
            # Label Placement
            target_idx = 0
            if label_col_idx < num_cols:
                target_idx = label_col_idx
            
            row_time[target_idx] = f"{line}"
            row_util[target_idx] = f"{line} Util(%)"
            
            for d_col in date_cols:
                t = col_line_totals.get((d_col, line), 0)
                row_time[d_col] = round(t, 1) if t > 0 else ""
                
                if t > 0:
                    util = (t / 480.0) * 100
                    row_util[d_col] = f"{util:.1f}%"
                else:
                    row_util[d_col] = "" 
            
            summ_rows.append(row_time)
            summ_rows.append(row_util)
            
        df_summ = pd.DataFrame(summ_rows, columns=df.columns)
        self.model_summary.set_dataframe(df_summ)

    def save_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                df = self.model_main.get_dataframe()
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"Saved to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def add_row(self):
        rows = self.model_main.rowCount()
        self.model_main.insertRows(rows, 1)

    def calculate_time(self):
        if not self.current_filepath:
            QMessageBox.warning(self, "Warning", "Please load or save an Excel file first.")
            return

        # 1. Update Summary with Current Data (Recalculate based on screen values)
        current_df = self.model_main.get_dataframe()
        split_col = getattr(self, 'split_col', 8) # default if not valid
        self.calc_summary(current_df, split_col)
        
        # Hide headers again for summary
        summ_df = self.model_summary.get_dataframe()
        new_cols = list(summ_df.columns)
        for i in range(split_col + 1):
            if i < len(new_cols):
                if i == split_col:
                    new_cols[i] = "생산라인"
                else:
                    new_cols[i] = ""
        summ_df.columns = new_cols
        self.model_summary.set_dataframe(summ_df)

        QMessageBox.information(self, "Success", "Production time and summary have been recalculated based on current table values.")



class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Date Range")
        self.layout = QVBoxLayout()
        
        # Start Date
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Start Date:"))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setDate(QDate.currentDate())
        h1.addWidget(self.start_date)
        self.layout.addLayout(h1)
        
        # End Date
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("End Date:  "))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setDate(QDate.currentDate())
        h2.addWidget(self.end_date)
        self.layout.addLayout(h2)
        
        # Buttons
        h3 = QHBoxLayout()
        btn_ok = QPushButton("OK")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        h3.addWidget(btn_ok)
        h3.addWidget(btn_cancel)
        self.layout.addLayout(h3)
        
        self.setLayout(self.layout)
        
    def get_date_range(self):
        s = self.start_date.date().toString("yyyy-MM-dd")
        e = self.end_date.date().toString("yyyy-MM-dd")
        return s, e

class OptimizationTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()
        
        # 1. File Inputs Group
        self.layout.addWidget(QLabel("<b>Input Files</b>"))
        self.input_layout = QVBoxLayout()
        
        self.bom_edit = self.create_file_input("BOM Folder:", "Input/BOM", mode='dir')
        self.common_edit = self.create_file_input("Common Material List:", "Input/common_material_list.csv")
        # self.schedule_edit = self.create_file_input("Schedule Excel:", "Input/smd_schedule.xlsx") # REMOVED
        
        self.layout.addLayout(self.input_layout)
        
        # 2. Options Group
        self.layout.addWidget(QLabel("<b>Options</b>"))
        
        # Priority (Row 1)
        row1_layout = QHBoxLayout()
        row1_layout.addWidget(QLabel("Priority Items:"))
        self.priority_edit = QLineEdit()
        self.priority_edit.setPlaceholderText("e.g. Item1, Item2 (comma separated)")
        row1_layout.addWidget(self.priority_edit)
        
        row1_layout.addWidget(QLabel("Layer Order:"))
        self.layer_edit = QLineEdit("TB") 
        self.layer_edit.setPlaceholderText("TB or BT")
        self.layer_edit.setFixedWidth(60)
        row1_layout.addWidget(self.layer_edit)
        self.layout.addLayout(row1_layout)
        
        # Manual Sequence (Row 2)
        self.layout.addWidget(QLabel("Manual Sequence (Optional):"))
        self.manual_edit = QTextEdit()
        self.manual_edit.setPlaceholderText("e.g. (ItemA, Top), (ItemB, Bottom) ...")
        self.manual_edit.setFixedHeight(60)
        self.layout.addWidget(self.manual_edit)
        
        # 3. Action
        self.btn_run = QPushButton("Run Optimization")
        self.btn_run.setFixedHeight(40)
        self.btn_run.clicked.connect(self.run_optimization)
        self.layout.addWidget(self.btn_run)
        
        # 4. Result View (Tabbed)
        self.layout.addWidget(QLabel("<b>Optimization Result</b>"))
        self.result_tabs = QTabWidget()
        self.layout.addWidget(self.result_tabs)
        
        self.btn_export = QPushButton("Export Result to Excel")
        self.btn_export.clicked.connect(self.export_result)
        self.layout.addWidget(self.btn_export)
        
        self.setLayout(self.layout)
        self.schedule_tab_ref = None

    def set_schedule_tab(self, tab):
        self.schedule_tab_ref = tab

    def create_file_input(self, label_text, default_path="", mode='file'):
        container = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0,0,0,0)
        
        lbl = QLabel(label_text)
        lbl.setFixedWidth(120)
        edit = QLineEdit(default_path)
        btn = QPushButton("Browse")
        btn.clicked.connect(lambda: self.browse_file(edit, mode))
        
        layout.addWidget(lbl)
        layout.addWidget(edit)
        layout.addWidget(btn)
        container.setLayout(layout)
        self.input_layout.addWidget(container)
        return edit

    def browse_file(self, line_edit, mode='file'):
        if mode == 'dir':
            path = QFileDialog.getExistingDirectory(self, "Select Directory", "")
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)")
        
        if path:
            line_edit.setText(path)

    def run_optimization(self):
        bom_folder = self.bom_edit.text()
        common_path = self.common_edit.text()
        
        if not all([os.path.exists(p) for p in [bom_folder, common_path]]):
            QMessageBox.warning(self, "Error", "Please verify all input paths exist.")
            return

        if not self.schedule_tab_ref:
             QMessageBox.warning(self, "Error", "Internal Error: Schedule Tab link missing.")
             return

        try:
            # Step 1: Prompt for Date Range
            dlg = DateRangeDialog(self)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
            
            start_str, end_str = dlg.get_date_range()
            
            # Generate Date List
            start_date = pd.to_datetime(start_str)
            end_date = pd.to_datetime(end_str)
            
            if start_date > end_date:
                QMessageBox.warning(self, "Error", "Start date cannot be after end date.")
                return
                
            date_list = pd.date_range(start=start_date, end=end_date).strftime('%Y-%m-%d').tolist()
            
            # Clear previous results
            self.result_tabs.clear()
            base_input_dir = "Input"
            
            processed_count = 0
            
            for date_str in date_list:
                print(f"Processing Date: {date_str}")
                
                # Step 2: Get Data from Schedule Tab
                items, error_msg = self.schedule_tab_ref.get_production_data(date_str)
                if error_msg or not items:
                    print(f"Skipping {date_str}: {error_msg or 'No items'}")
                    continue
                
                # Save item_list.txt
                item_list_dst = os.path.join(base_input_dir, "item_list.txt")
                with open(item_list_dst, 'w', encoding='utf-8-sig') as f:
                    f.write("Item_Code,T_B,Qty,Prod_Time\n")
                    for item in items:
                         tb_char = 'T' if item['Layer'] == 'Top' else 'B'
                         f.write(f"{item['Item_Code']},{tb_char},{item['Qty']},{item['Prod_Time']}\n")

                # Step 3: Build Dynamic BOM
                required_items = set()
                with open(item_list_dst, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        if 'Item_Code' in row:
                            required_items.add(row['Item_Code'].strip())
                
                merged_bom_path = os.path.join(base_input_dir, "BOM.txt")
                with open(merged_bom_path, 'w', encoding='utf-8-sig', newline='') as f_out:
                    writer = None
                    for item_code in required_items:
                        bom_file = os.path.join(bom_folder, f"{item_code}.txt")
                        if os.path.exists(bom_file):
                            with open(bom_file, 'r', encoding='utf-8-sig') as f_in:
                                lines = f_in.readlines()
                                if lines:
                                    if writer is None:
                                        f_out.write(lines[0]) 
                                        writer = True
                                    for line in lines[1:]:
                                        if line.strip():
                                            f_out.write(line)

                # Step 4 & 5: Run Optimization Scripts
                
                # CLEANUP: Remove stale output from previous iteration to avoid Duplicate/Wrong data
                result_path = "Output/optimization_sequence.csv"
                if os.path.exists(result_path):
                    os.remove(result_path)

                cmd_plan = [sys.executable, "optimize_plan.py"]
                subprocess.run(cmd_plan, check=True)
                
                cmd_seq = [sys.executable, "optimize_sequence.py"]
                if self.priority_edit.text():
                    cmd_seq.extend(["--priority", self.priority_edit.text()])
                if self.layer_edit.text():
                    cmd_seq.extend(["--layer", self.layer_edit.text()])
                manual_text = self.manual_edit.toPlainText().replace('\n', ' ').strip()
                if manual_text:
                    cmd_seq.extend(["--manual", manual_text])
                    
                subprocess.run(cmd_seq, check=True)
                
                # Load Result
                df_res = pd.DataFrame()
                if os.path.exists(result_path):
                    df_res = pd.read_csv(result_path)
                else:
                    # If no result generated (e.g. 0 items), create empty DF with standard columns
                    cols = ['Index', 'Item_Code', 'Layer', 'Qty', 'Prod_Time', 'Total_Count', 
                            'Common_Count', 'Individual_Count', 'Transition_Shared_Count', 'Selection_Reason']
                    df_res = pd.DataFrame(columns=cols)
                    
                # Create Tab
                tab = QWidget()
                tab_layout = QVBoxLayout()
                table = QTableView()
                model = PandasModel(df_res)
                table.setModel(model)
                tab_layout.addWidget(table)
                tab.setLayout(tab_layout)
                
                # Store model in tab for export
                table.setProperty("model_ref", model) 
                
                self.result_tabs.addTab(tab, date_str)
                processed_count += 1
            
            if processed_count == 0:
                QMessageBox.warning(self, "Info", "No data processed for the selected range.")
            else:
                QMessageBox.information(self, "Done", f"Optimization Complete! Processed {processed_count} days.")
            
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Error", f"Script failed: {e}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", str(e))
            


    def export_result(self):
        curr_idx = self.result_tabs.currentIndex()
        if curr_idx == -1: return
        
        # Get Current Tab Data
        # We stored reference or can find child table
        tab = self.result_tabs.widget(curr_idx)
        table = tab.findChild(QTableView)
        model = table.model() # PandasModel
        
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Result Excel", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                df = model.get_dataframe()
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"Saved to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Production Optimization & Schedule Manager")
        self.resize(1200, 800)
        
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Tab 1
        self.tab_sm_verify = SMDVerificationTab()
        self.tabs.addTab(self.tab_sm_verify, "SMD Data Verification")

        # Tab 2
        self.tab_schedule = ScheduleTab()
        self.tabs.addTab(self.tab_schedule, "Schedule Management")
        
        # Tab 3
        self.tab_optimize = OptimizationTab()
        self.tabs.addTab(self.tab_optimize, "Production Optimization")
        
        # Link Tabs
        self.tab_optimize.set_schedule_tab(self.tab_schedule)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
