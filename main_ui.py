import sys
import os
import subprocess
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTableView, QFileDialog, QTabWidget, QLabel, 
                             QLineEdit, QMessageBox, QHeaderView, QAbstractItemView,
                             QInputDialog, QDialog, QTextEdit, QTableWidget, QTableWidgetItem)
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt6.QtGui import QColor, QFont
import shutil

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

        # 1. Update Summary with Current Data
        current_df = self.model_main.get_dataframe()
        split_col = getattr(self, 'split_col', 8) # default if not valid
        self.calc_summary(current_df, split_col)
        
        # Hide headers again for summary
        summ_df = self.model_summary.get_dataframe()
        new_cols = list(summ_df.columns)
        for i in range(split_col + 1):
            if i < len(new_cols): new_cols[i] = ""
        summ_df.columns = new_cols
        self.model_summary.set_dataframe(summ_df)
            
        date_str, ok = QInputDialog.getText(self, "Select Date", "Enter Date (YYYY-MM-DD) matching column header:", text="2025-12-26")
        if not ok: return
        
        setup_args = []
        for line, le in self.setup_inputs.items():
            val = le.text().strip()
            if val:
                setup_args.append(f"{line}:{val}")
        setup_arg_str = ",".join(setup_args)
        
        # Save temp file for subprocess
        try:
            temp_path = os.path.join(os.path.dirname(self.current_filepath), "~temp_calc.xlsx")
            current_df.to_excel(temp_path, index=False)
            
            cmd = [
                sys.executable, 
                r"D:\Develoment\ProductOptimize\calculate_schedule.py",
                "--file", temp_path,
                "--date", date_str,
                "--setup-times", setup_arg_str
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='cp949', errors='replace')
            output_text = result.stdout
            if result.stderr:
                output_text += "\n[STDERR]\n" + result.stderr
                
            # Cleanup
            if os.path.exists(temp_path):
                os.remove(temp_path)

            dlg = QDialog(self)
            dlg.setWindowTitle("Production Time Calculation Result")
            dlg.resize(700, 500)
            layout = QVBoxLayout()
            
            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            text_edit.setFont(QFont("Courier New", 10))
            text_edit.setText(output_text)
            
            layout.addWidget(text_edit)
            dlg.setLayout(layout)
            dlg.exec()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


class OptimizationTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()
        
        # 1. File Inputs Group
        self.layout.addWidget(QLabel("<b>Input Files</b>"))
        self.input_layout = QVBoxLayout()
        
        self.bom_edit = self.create_file_input("BOM File:", "D:\\Develoment\\ProductOptimize\\Input\\BOM.txt")
        self.common_edit = self.create_file_input("Common Material List:", "D:\\Develoment\\ProductOptimize\\Input\\common_material_list.csv")
        self.schedule_edit = self.create_file_input("Schedule Excel:", "D:\\Develoment\\ProductOptimize\\Input\\smd_schedule.xlsx")
        
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
        
        # 4. Result View
        self.layout.addWidget(QLabel("<b>Optimization Result</b>"))
        self.result_table = QTableView()
        self.result_model = PandasModel(pd.DataFrame())
        self.result_table.setModel(self.result_model)
        self.layout.addWidget(self.result_table)
        
        self.btn_export = QPushButton("Export Result to Excel")
        self.btn_export.clicked.connect(self.export_result)
        self.layout.addWidget(self.btn_export)
        
        self.setLayout(self.layout)

    def create_file_input(self, label_text, default_path=""):
        container = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0,0,0,0)
        
        lbl = QLabel(label_text)
        lbl.setFixedWidth(120)
        edit = QLineEdit(default_path)
        btn = QPushButton("Browse")
        btn.clicked.connect(lambda: self.browse_file(edit))
        
        layout.addWidget(lbl)
        layout.addWidget(edit)
        layout.addWidget(btn)
        container.setLayout(layout)
        self.input_layout.addWidget(container)
        return edit

    def browse_file(self, line_edit):
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)")
        if path:
            line_edit.setText(path)

    def run_optimization(self):
        bom_path = self.bom_edit.text()
        common_path = self.common_edit.text()
        schedule_path = self.schedule_edit.text()
        
        if not all([os.path.exists(p) for p in [bom_path, common_path, schedule_path]]):
            QMessageBox.warning(self, "Error", "Please verify all input file paths exist.")
            return

        try:
            base_input_dir = r"D:\Develoment\ProductOptimize\Input"
            
            def safe_copy(src, dst):
                if os.path.abspath(src) != os.path.abspath(dst):
                    shutil.copy(src, dst)
                    print(f"Copied {src} to {dst}")
                else:
                    print(f"Skipping copy: {src} is same as destination.")

            safe_copy(bom_path, os.path.join(base_input_dir, "BOM.txt"))
            safe_copy(common_path, os.path.join(base_input_dir, "common_material_list.csv"))
            
            # Step 1: Optimize Plan (Analyze BOM)
            cmd_plan = [sys.executable, r"D:\Develoment\ProductOptimize\optimize_plan.py"]
            subprocess.run(cmd_plan, check=True)
            
            # Step 2: Calculate Schedule (Get details for specific date)
            # Need date. Prompt user.
            date_str, ok = QInputDialog.getText(self, "Date Selection", "Enter Production Date (matching Excel header):", text="2025-12-26")
            if not ok: return
            
            cmd_schedule = [
                sys.executable, 
                r"D:\Develoment\ProductOptimize\calculate_schedule.py",
                "--file", schedule_path,
                "--date", date_str
            ]
            subprocess.run(cmd_schedule, check=True)
            
            # Move result to item_list.txt for optimizer
            safe_copy(
                os.path.join(base_input_dir, "item_list_from_excel.txt"),
                os.path.join(base_input_dir, "item_list.txt")
            )
            
            # Step 3: Run Sequence Optimization
            cmd_seq = [sys.executable, r"D:\Develoment\ProductOptimize\optimize_sequence.py"]
            
            if self.priority_edit.text():
                cmd_seq.extend(["--priority", self.priority_edit.text()])
            
            if self.layer_edit.text():
                cmd_seq.extend(["--layer", self.layer_edit.text()])

            manual_text = self.manual_edit.toPlainText().replace('\n', ' ').strip()
            if manual_text:
                cmd_seq.extend(["--manual", manual_text])
                
            subprocess.run(cmd_seq, check=True)
            
            # Load Result
            result_path = r"D:\Develoment\ProductOptimize\Output\optimization_sequence.csv"
            if os.path.exists(result_path):
                df_res = pd.read_csv(result_path)
                self.result_model.set_dataframe(df_res)
                QMessageBox.information(self, "Done", "Optimization Complete!")
            
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Error", f"Script failed: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def export_result(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Result Excel", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                df = self.result_model.get_dataframe()
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
        self.tab_schedule = ScheduleTab()
        self.tabs.addTab(self.tab_schedule, "Schedule Management")
        
        # Tab 2
        self.tab_optimize = OptimizationTab()
        self.tabs.addTab(self.tab_optimize, "Production Optimization")
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
