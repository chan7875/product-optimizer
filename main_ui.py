import sys
import os
import subprocess
import pandas as pd
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTableView, QFileDialog, QTabWidget, QLabel, 
                             QLineEdit, QMessageBox, QHeaderView, QAbstractItemView,
                             QInputDialog, QDialog, QTextEdit, QTableWidget, QTableWidgetItem,
                             QDateEdit, QSplitter, QTreeWidget, QTreeWidgetItem, QStackedWidget, QMenu, QStackedLayout,
                             QGraphicsView, QGraphicsScene)
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QDate, QUrl, QEvent, QPoint, QPointF, QRectF
from PyQt6.QtGui import QColor, QFont, QCursor, QKeySequence, QWheelEvent, QPen, QBrush, QPainterPath, QPolygonF, QTransform
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineSettings, QWebEnginePage
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


class NeutralFileParser:
    def __init__(self):
        self.board_outline = []
        self.geometries = {}
        self.components = []
        
    def parse(self, file_path):
        self.board_outline = []
        self.geometries = {}
        self.components = []
        
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
        except Exception as e:
            print(f"Error reading file: {e}")
            return None
        
        # Join continuation lines (ending with -)
        joined_lines = []
        current_line = ""
        for line in lines:
            stripped = line.rstrip()
            if stripped.endswith('-'):
                current_line += stripped[:-1] + " "  # Remove '-' and add space
            else:
                current_line += stripped
                joined_lines.append(current_line)
                current_line = ""
        if current_line:
            joined_lines.append(current_line)
        
        current_section = None
        current_geom_name = None
        
        for line in joined_lines:
            line = line.strip()
            
            # Section Detection
            if 'Attribute Information' in line:
                current_section = 'ATTR'
                continue
            elif 'Geometry Information' in line:
                current_section = 'GEOM'
                continue
            elif 'Component Information' in line:
                current_section = 'COMP'
                continue
            elif line.startswith('#'):
                continue  # Skip comment/delimiter lines
            
            # Parse Attribute Section (Board Outline)
            if current_section == 'ATTR' and line.startswith("B_ATTR") and "'BOARD_AREA'" in line:
                coords = self._extract_coords(line)
                self.board_outline = coords
            
            # Parse Geometry Section
            if current_section == 'GEOM':
                if line.startswith('GEOM '):
                    parts = line.split()
                    if len(parts) >= 2:
                        current_geom_name = parts[1]
                        self.geometries[current_geom_name] = []
                elif line.startswith("G_ATTR") and "'COMPONENT_PLACEMENT_OUTLINE'" in line and current_geom_name:
                    coords = self._extract_coords(line)
                    self.geometries[current_geom_name] = coords
            
            # Parse Component Section
            if current_section == 'COMP':
                if line.startswith('COMP '):
                    parts = line.split()
                    if len(parts) >= 9:
                        try:
                            comp = {
                                'ref': parts[1],
                                'part_no': parts[2],
                                'name': parts[3],
                                'geom_name': parts[4],
                                'x': float(parts[5]),
                                'y': float(parts[6]),
                                'layer': int(parts[7]),
                                'rotation': float(parts[8]),
                                'properties': {}  # Will be filled by C_PROP
                            }
                            self.components.append(comp)
                        except (ValueError, IndexError) as e:
                            print(f"Error parsing COMP line: {e}")
                elif line.startswith('C_PROP ') and self.components:
                    # Parse C_PROP and add to last component
                    props = self._parse_c_prop(line)
                    self.components[-1]['properties'].update(props)
        
        print(f"Parsed: Board Outline={len(self.board_outline)} pts, Geometries={len(self.geometries)}, Components={len(self.components)}")
        return {
            'board_outline': self.board_outline,
            'geometries': self.geometries,
            'components': self.components
        }
    
    def _extract_coords(self, line):
        import re
        coords = []
        # Find position after last quote (coordinates start after 'Layer1' or '' )
        last_quote = line.rfind("'")
        if last_quote != -1:
            coord_part = line[last_quote+1:]
        else:
            coord_part = line
        
        # Extract all floating point numbers
        numbers = re.findall(r'[-+]?\d+\.?\d*', coord_part)
        
        # Pair them as X, Y coordinates
        for j in range(0, len(numbers) - 1, 2):
            try:
                x = float(numbers[j])
                y = float(numbers[j+1])
                coords.append((x, y))
            except ValueError:
                pass
        return coords
    
    def _parse_c_prop(self, line):
        """Parse C_PROP line to extract name/value pairs from (NAME,"VALUE") format"""
        import re
        props = {}
        # Remove 'C_PROP ' prefix
        content = line[7:] if line.startswith('C_PROP ') else line
        
        # Pattern to match (NAME,"VALUE") or (NAME,VALUE)
        pattern = r'\(([^,]+),\"?([^\")]+)\"?\)'
        matches = re.findall(pattern, content)
        
        for name, value in matches:
            props[name.strip()] = value.strip()
        
        return props

class ZoomableGraphicsView(QGraphicsView):
    """QGraphicsView with mouse wheel zoom support"""
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        
    def wheelEvent(self, event):
        # Zoom Factor
        zoom_in_factor = 1.25
        zoom_out_factor = 1 / zoom_in_factor
        
        if event.angleDelta().y() > 0:
            zoom_factor = zoom_in_factor
        else:
            zoom_factor = zoom_out_factor
            
        self.scale(zoom_factor, zoom_factor)

class CADViewerWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(2)
        
        # Toolbar for rotation buttons
        self.toolbar = QHBoxLayout()
        self.toolbar.setContentsMargins(5, 5, 5, 5)
        self.main_layout.addLayout(self.toolbar)
        
        self.toolbar.addStretch()
        self.btn_0 = QPushButton("0°")
        self.btn_90 = QPushButton("90°")
        self.btn_180 = QPushButton("180°")
        self.btn_270 = QPushButton("270°")
        
        for btn, angle in [(self.btn_0, 0), (self.btn_90, 90), (self.btn_180, 180), (self.btn_270, 270)]:
            btn.setFixedWidth(50)
            btn.clicked.connect(lambda checked, a=angle: self.rotate_view(a))
            self.toolbar.addWidget(btn)
        
        # Separator
        self.toolbar.addSpacing(20)
        
        # Zoom Fit Button
        self.btn_fit = QPushButton("Zoom Fit")
        self.btn_fit.setFixedWidth(100)
        self.btn_fit.clicked.connect(self.zoom_fit)
        self.toolbar.addWidget(self.btn_fit)
        
        # Separator
        self.toolbar.addSpacing(10)
        
        # Search Component
        self.lbl_search = QLabel("Find:")
        self.toolbar.addWidget(self.lbl_search)
        self.input_search = QLineEdit()
        self.input_search.setPlaceholderText("Ref...")
        self.input_search.setFixedWidth(80)
        self.input_search.returnPressed.connect(self.find_component)
        self.toolbar.addWidget(self.input_search)
        self.btn_search = QPushButton("Search")
        self.btn_search.clicked.connect(self.find_component)
        self.toolbar.addWidget(self.btn_search)
        
        self.toolbar.addStretch()
        
        # Horizontal Splitter: CAD View (left) + Property Table (right)
        self.content_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_layout.addWidget(self.content_splitter)
        
        # Left: Tab Widget for Top/Bottom
        self.tabs = QTabWidget()
        self.content_splitter.addWidget(self.tabs)
        
        # Top Layer
        self.scene_top = QGraphicsScene()
        self.view_top = ZoomableGraphicsView(self.scene_top)
        self.view_top.setMouseTracking(True)
        self.tabs.addTab(self.view_top, "Top")
        
        # Bottom Layer
        self.scene_bottom = QGraphicsScene()
        self.view_bottom = ZoomableGraphicsView(self.scene_bottom)
        self.view_bottom.setMouseTracking(True)
        self.tabs.addTab(self.view_bottom, "Bottom")
        
        # Right: Property Table
        self.prop_table = QTableWidget()
        self.prop_table.setColumnCount(2)
        self.prop_table.setHorizontalHeaderLabels(["Property", "Value"])
        self.prop_table.horizontalHeader().setStretchLastSection(True)
        self.prop_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.prop_table.setMinimumWidth(200)
        self.content_splitter.addWidget(self.prop_table)
        
        # Set splitter sizes (80% CAD, 20% property)
        self.content_splitter.setSizes([800, 200])
        
        # Store component bounds for search
        self.component_bounds = {}  # ref -> (layer, QRectF)
        self.component_items = {}   # ref -> QGraphicsPathItem (for highlighting)
        self.component_data = {}    # ref -> component dict (with properties)
        self.highlighted_ref = None  # Currently highlighted component ref
        self.parser = NeutralFileParser()
        self.current_rotation = 0
        
        # Connect mouse click events on views
        self.view_top.viewport().installEventFilter(self)
        self.view_bottom.viewport().installEventFilter(self)
    
    def eventFilter(self, obj, event):
        """Handle mouse click on graphics views to select components"""
        from PyQt6.QtCore import QEvent
        if event.type() == QEvent.Type.MouseButtonPress:
            if obj == self.view_top.viewport():
                self._handle_view_click(self.view_top, self.scene_top, 1, event.pos())
            elif obj == self.view_bottom.viewport():
                self._handle_view_click(self.view_bottom, self.scene_bottom, 2, event.pos())
        return super().eventFilter(obj, event)
    
    def _handle_view_click(self, view, scene, layer, pos):
        """Find clicked component and show its properties"""
        scene_pos = view.mapToScene(pos)
        
        # Find which component was clicked
        for ref, (comp_layer, bounds) in self.component_bounds.items():
            if comp_layer == layer and bounds.contains(scene_pos):
                self._show_component_properties(ref)
                return
    
    def _show_component_properties(self, ref):
        """Display component properties in the property table"""
        self.prop_table.setRowCount(0)
        
        if ref not in self.component_data:
            return
            
        comp = self.component_data[ref]
        
        # Add basic component info
        basic_props = [
            ("Reference", comp.get('ref', '')),
            ("Part Number", comp.get('part_no', '')),
            ("Name", comp.get('name', '')),
            ("Geometry", comp.get('geom_name', '')),
            ("X", str(comp.get('x', ''))),
            ("Y", str(comp.get('y', ''))),
            ("Layer", "Top" if comp.get('layer') == 1 else "Bottom"),
            ("Rotation", str(comp.get('rotation', '')))
        ]
        
        # Add C_PROP properties
        properties = comp.get('properties', {})
        
        all_props = basic_props + list(properties.items())
        
        self.prop_table.setRowCount(len(all_props))
        for i, (name, value) in enumerate(all_props):
            self.prop_table.setItem(i, 0, QTableWidgetItem(str(name)))
            self.prop_table.setItem(i, 1, QTableWidgetItem(str(value)))
        
        self.prop_table.resizeColumnsToContents()
        
    def rotate_view(self, angle):
        # Calculate delta rotation
        delta = angle - self.current_rotation
        self.current_rotation = angle
        
        self.view_top.rotate(delta)
        self.view_bottom.rotate(delta)
    
    def zoom_fit(self):
        """Fit the current view to show entire scene"""
        current_view = self.view_top if self.tabs.currentIndex() == 0 else self.view_bottom
        current_scene = self.scene_top if self.tabs.currentIndex() == 0 else self.scene_bottom
        current_view.fitInView(current_scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
        
    def find_component(self):
        ref = self.input_search.text().strip().upper()
        if not ref:
            return
        
        # Reset previous highlight
        if self.highlighted_ref and self.highlighted_ref in self.component_items:
            prev_item = self.component_items[self.highlighted_ref]
            prev_item.setBrush(QBrush(QColor(255, 255, 255)))  # White
        
        # Search in component_bounds (case-insensitive)
        found = None
        found_ref = None
        for comp_ref, (layer, bounds) in self.component_bounds.items():
            if comp_ref.upper() == ref:
                found = (layer, bounds)
                found_ref = comp_ref
                break
        
        if found:
            layer, bounds = found
            
            # Highlight the found component
            if found_ref in self.component_items:
                item = self.component_items[found_ref]
                item.setBrush(QBrush(QColor(220, 220, 220)))  # Light gray
                self.highlighted_ref = found_ref
            
            # Show properties in property grid
            self._show_component_properties(found_ref)
            
            # Switch to correct tab
            if layer == 1:
                self.tabs.setCurrentIndex(0)
                view = self.view_top
            else:
                self.tabs.setCurrentIndex(1)
                view = self.view_bottom
            
            # Create expanded rect (4x size centered on component)
            center = bounds.center()
            expanded_width = bounds.width() * 4
            expanded_height = bounds.height() * 4
            expanded_rect = QRectF(
                center.x() - expanded_width/2,
                center.y() - expanded_height/2,
                expanded_width,
                expanded_height
            )
            
            # Zoom to the expanded rect
            view.fitInView(expanded_rect, Qt.AspectRatioMode.KeepAspectRatio)
        else:
            self.highlighted_ref = None
            self.prop_table.setRowCount(0)  # Clear properties
            print(f"Component '{ref}' not found")
        
    def load_neutral_file(self, file_path):
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return
            
        # Reset
        self.current_rotation = 0
        self.view_top.resetTransform()
        self.view_bottom.resetTransform()
        self.component_bounds.clear()
        self.component_items.clear()
        self.component_data.clear()
        self.highlighted_ref = None
        self.prop_table.setRowCount(0)
            
        data = self.parser.parse(file_path)
        if data:
            self._draw_cad(data)
            
    def _draw_cad(self, data):
        self.scene_top.clear()
        self.scene_bottom.clear()
        
        board_outline = data['board_outline']
        geometries = data['geometries']
        components = data['components']
        
        # Store component data for property display
        for comp in components:
            self.component_data[comp['ref']] = comp
        
        # Draw board outline on both scenes
        for scene in [self.scene_top, self.scene_bottom]:
            self._draw_board_outline(scene, board_outline)
        
        # Draw components by layer (1=Top, 2=Bottom)
        top_comps = [c for c in components if c['layer'] == 1]
        bottom_comps = [c for c in components if c['layer'] == 2]
        
        self._draw_components(self.scene_top, top_comps, geometries)
        self._draw_components(self.scene_bottom, bottom_comps, geometries)
        
        # Fit views
        self.view_top.fitInView(self.scene_top.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.view_bottom.fitInView(self.scene_bottom.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
        
    def _draw_board_outline(self, scene, points):
        if not points:
            return
        pen = QPen(QColor(0, 0, 255))
        pen.setWidth(2)
        for i in range(len(points)):
            x1, y1 = points[i]
            x2, y2 = points[(i + 1) % len(points)]
            scene.addLine(x1, -y1, x2, -y2, pen)
            
    def _draw_components(self, scene, components, geometries):
        pen = QPen(QColor(0, 128, 0))
        pen.setWidthF(1.0)  # Width in pixels
        pen.setCosmetic(True)  # Cosmetic pen - width stays constant regardless of zoom
        brush = QBrush(QColor(255, 255, 255))  # White fill
        
        for comp in components:
            geom_name = comp['geom_name']
            geom_points = geometries.get(geom_name, [])
            comp_x = comp['x']
            comp_y = -comp['y']  # Flip Y
            comp_layer = comp['layer']
            comp_ref = comp['ref']
            
            # Calculate component bounding box for text scaling
            comp_width = 2.0
            comp_height = 2.0
            comp_bounds = QRectF(comp_x - 1, comp_y - 1, 2, 2)  # Default bounds
            path_item = None
            
            if geom_points:
                path = QPainterPath()
                if len(geom_points) > 0:
                    path.moveTo(geom_points[0][0], -geom_points[0][1])
                    for pt in geom_points[1:]:
                        path.lineTo(pt[0], -pt[1])
                    path.closeSubpath()
                
                # Get bounding rect of geometry for text scaling
                geom_rect = path.boundingRect()
                comp_width = geom_rect.width()
                comp_height = geom_rect.height()
                
                transform = QTransform()
                transform.translate(comp_x, comp_y)
                transform.rotate(-comp['rotation'])
                
                transformed_path = transform.map(path)
                path_item = scene.addPath(transformed_path, pen, brush)
                
                # Store transformed bounds for search
                comp_bounds = transformed_path.boundingRect()
            else:
                rect_size = 2
                path_item = scene.addRect(comp_x - rect_size/2, comp_y - rect_size/2, 
                              rect_size, rect_size, pen, brush)
                comp_width = rect_size
                comp_height = rect_size
                comp_bounds = QRectF(comp_x - rect_size/2, comp_y - rect_size/2, rect_size, rect_size)
            
            # Store bounds and items for component search/highlight
            self.component_bounds[comp_ref] = (comp_layer, comp_bounds)
            if path_item:
                self.component_items[comp_ref] = path_item
            
            # Add Reference Text - use monospace font for better visibility
            # Use QGraphicsSimpleTextItem for better scaling control (no margins)
            font = QFont("Consolas")
            font.setPointSizeF(100)  # Use large base size for better resolution when scaled down
            font.setBold(True)      # Make it bold
            
            text_item = scene.addSimpleText(comp['ref'], font)
            text_item.setBrush(QBrush(QColor(0, 0, 0)))
            text_item.setPen(QPen(Qt.PenStyle.NoPen))  # No outline for text
            
            # Check for vertical rotation
            is_vertical = False
            rot = comp['rotation']
            # Normalize rotation check
            if abs(rot - 90) < 1.0 or abs(rot - 270) < 1.0:
                is_vertical = True
            
            # Get text bounding rect
            text_rect = text_item.boundingRect()
            
            # Calculate scale to fit text within component bounds (90% of size)
            target_width = comp_width * 0.9
            target_height = comp_height * 0.9
            
            # If vertical, we need to swap target width/height because we will rotate text
            if is_vertical:
                # Text Height (visual width) should check against Comp Width
                # Text Width (visual height) should check against Comp Height
                # Since we want to fit text inside the box:
                # Local Text Width fits into Local Comp Height (Visual Height)
                # Local Text Height fits into Local Comp Width (Visual Width)
                # target_width is usually X-axis. 
                # If we rotate text 90, its X-axis aligns with Comp Y-axis.
                # So we swap targets for scaling calculation.
                scale_x = target_height / text_rect.width()
                scale_y = target_width / text_rect.height()
            else:
                scale_x = target_width / text_rect.width()
                scale_y = target_height / text_rect.height()
                
            if text_rect.width() > 0 and text_rect.height() > 0:
                # Use smaller scale to maintain aspect ratio
                scale = min(scale_x, scale_y)
                text_item.setScale(scale)
            
            # Reset position to 0,0 for mapping calculation
            text_item.setPos(0, 0)
            
            # Apply Rotation if vertical
            if is_vertical:
                center = text_rect.center()
                text_item.setTransformOriginPoint(center)
                text_item.setRotation(-90)
                
            # Robust Centering Logic:
            # 1. Get where the center of the text is currently mapped to in the scene (with Pos=0,0)
            rect_center = text_rect.center()
            current_scene_center = text_item.mapToScene(rect_center)
            
            # 2. Calculate the difference between where it is and where we want it (comp_x, comp_y)
            target_pos = QPointF(comp_x, comp_y)
            offset = target_pos - current_scene_center
            
            # 3. Apply that offset to the item's position
            text_item.setPos(offset)
    
    def mark_only_cad_components(self, only_cad_refs):
        """Mark components that are Only CAD (not in BOM) with a red X"""
        if not only_cad_refs:
            return
            
        pen = QPen(QColor(255, 0, 0))  # Red
        pen.setWidthF(2.0)
        pen.setCosmetic(True)  # Width stays constant regardless of zoom
        
        for ref in only_cad_refs:
            ref = ref.strip()
            if ref in self.component_bounds:
                layer, bounds = self.component_bounds[ref]
                
                # Select the appropriate scene
                if layer == 1:
                    scene = self.scene_top
                else:
                    scene = self.scene_bottom
                
                # Draw X from corner to corner of bounding box
                x1, y1 = bounds.left(), bounds.top()
                x2, y2 = bounds.right(), bounds.bottom()
                
                # Diagonal line 1: top-left to bottom-right
                scene.addLine(x1, y1, x2, y2, pen)
                # Diagonal line 2: top-right to bottom-left
                scene.addLine(x2, y1, x1, y2, pen)


# Import logic from existing scripts
import calculate_schedule


class HandToolOverlay(QWidget):
    def __init__(self, parent=None, web_view=None):
        super().__init__(parent)
        self.web_view = web_view
        self.last_pos = None
        self.setCursor(Qt.CursorShape.OpenHandCursor)
        # Transparent background? No, let's use semi-transparent for debugging
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("background-color: rgba(0, 255, 0, 50);") # Debug Green
        self.hide() 
        
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.last_pos = event.pos()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
            print("Overlay: Mouse Press")
            event.accept()

    def mouseMoveEvent(self, event):
        if self.last_pos:
            delta = event.pos() - self.last_pos
            self.last_pos = event.pos()
            print(f"Overlay: Drag delta={delta}, Sending Native Wheel")
            if self.web_view:
                # Create native Qt wheel event
                # angleDelta is in 1/8 degree units; multiply by 8 for responsiveness
                angle_delta = QPoint(delta.x() * 8, delta.y() * 8)
                pixel_delta = QPoint(delta.x(), delta.y())
                
                # Get center of view for event position
                center = self.web_view.rect().center()
                
                wheel_event = QWheelEvent(
                    QPointF(center),
                    QPointF(self.web_view.mapToGlobal(center)),
                    pixel_delta,
                    angle_delta,
                    Qt.MouseButton.NoButton,
                    Qt.KeyboardModifier.NoModifier,
                    Qt.ScrollPhase.NoScrollPhase,
                    False
                )
                
                # Send to focusProxy (the actual render widget) or the view
                target = self.web_view.focusProxy() or self.web_view
                QApplication.sendEvent(target, wheel_event)
            event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.last_pos = None
            self.setCursor(Qt.CursorShape.OpenHandCursor)
            event.accept()

class PDFWebView(QWebEngineView):
    pass

class PDFViewerWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0,0,0,0)
        self.layout.setSpacing(2)
        
        # Toolbar
        self.toolbar_layout = QHBoxLayout()
        self.toolbar_layout.setContentsMargins(5, 5, 5, 5)
        self.layout.addLayout(self.toolbar_layout)
        
        # Hand Tool Toggle
        self.btn_hand = QPushButton("Hand Tool")
        self.btn_hand.setCheckable(True)
        self.btn_hand.toggled.connect(self.toggle_hand_mode)
        self.toolbar_layout.addWidget(self.btn_hand)
        
        self.toolbar_layout.addStretch()
        
        # Find Bar
        self.lbl_find = QLabel("Find:")
        self.toolbar_layout.addWidget(self.lbl_find)
        
        self.input_find = QLineEdit()
        self.input_find.setPlaceholderText("Text...")
        self.input_find.setFixedWidth(150)
        self.input_find.returnPressed.connect(self.find_next)
        self.toolbar_layout.addWidget(self.input_find)
        
        self.btn_prev = QPushButton("Prev")
        self.btn_prev.clicked.connect(self.find_prev)
        self.toolbar_layout.addWidget(self.btn_prev)
        
        self.btn_next = QPushButton("Next")
        self.btn_next.clicked.connect(self.find_next)
        self.toolbar_layout.addWidget(self.btn_next)
        
        # Content Container with Stacked Layout (Overlay on top of View)
        self.content_container = QWidget()
        self.stack = QStackedLayout(self.content_container)
        self.stack.setStackingMode(QStackedLayout.StackingMode.StackAll)
        
        # WebView
        self.view = PDFWebView()
        self.stack.addWidget(self.view)
        
        # Overlay
        self.overlay = HandToolOverlay(self.content_container, self.view)
        self.stack.addWidget(self.overlay)
        
        self.layout.addWidget(self.content_container)
        
    def toggle_hand_mode(self, checked):
        if checked:
            self.overlay.show()
            self.overlay.raise_()
        else:
            self.overlay.hide()
        
    def find_next(self):
        text = self.input_find.text()
        if text:
            # Pass callback to handle zoom
            self.view.findText(text, resultCallback=self.on_find_result)
            
    def find_prev(self):
        text = self.input_find.text()
        if text:
            self.view.findText(text, QWebEnginePage.FindFlag.FindBackward, resultCallback=self.on_find_result)

    def on_find_result(self, found):
        if found:
            # Zoom logic
            current = self.view.zoomFactor()
            if current < 2.0:
                self.view.setZoomFactor(2.0)
        else:
            # Optional: Feedback for not found
            print("Text not found")
            
    def setUrl(self, url):
        self.view.setUrl(url)
        
    def settings(self):
        return self.view.settings()

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

        self.lbl_path = QLabel("No folder selected")
        
        self.btn_run_smd = QPushButton("SMD Pro 실행")
        self.btn_run_smd.clicked.connect(self.run_smd_pro)
        
        top_bar.addWidget(self.btn_load_folder)
        top_bar.addWidget(self.lbl_path)
        top_bar.addWidget(self.btn_run_smd)
        top_bar.addStretch()
        
        self.layout.addWidget(self.top_container)
        
        # Main Horizontal Splitter
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: Tree Widget (Filter/Navigation)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("SMD Code / PCB Code")
        self.tree.setFixedWidth(250)
        self.splitter.addWidget(self.tree)
        
        # Right Container (Vertical Splitter)
        self.right_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Right Top: Table Widget (Details)
        self.table = QTableWidget()
        cols = ["Check", "PCB Code", "Rev", "SMD Code", "Model Name", "PCB Size", 
                "Neutral File", "Gerber File", "BOM File", "BOM Count", "WorkSpec", "Path"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        # Set selection color to light gray
        self.table.setStyleSheet("QTableWidget::item:selected { background-color: rgb(220, 220, 220); color: black; }")
        self.right_splitter.addWidget(self.table)
        
        # Right Bottom: Stacked Widget (Report vs PDF)
        self.bottom_stack = QStackedWidget()
        self.report_tabs = QTabWidget()
        self.web_view = PDFViewerWidget()
        # Enable Settings for PDF (via proxy)
        self.web_view.settings().setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, True)
        self.web_view.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        
        self.bottom_stack.addWidget(self.report_tabs)
        self.bottom_stack.addWidget(self.web_view)
        
        # CAD Viewer for Neutral File
        self.cad_viewer = CADViewerWidget()
        self.bottom_stack.addWidget(self.cad_viewer)
        
        self.right_splitter.addWidget(self.bottom_stack)
        
        # Set Right Splitter Stretch (30:70)
        self.right_splitter.setStretchFactor(0, 3)
        self.right_splitter.setStretchFactor(1, 7)
        # Enforce initial size ratio
        self.right_splitter.setSizes([300, 700])
        
        self.splitter.addWidget(self.right_splitter)
        
        # Set Main Splitter Stretch (Tree:Right => FixedWidth:Expand)
        self.splitter.setStretchFactor(1, 1) 
        
        self.layout.addWidget(self.splitter)
        self.setLayout(self.layout)
        
        self.json_data_list = []
        
        self.tree.itemClicked.connect(self.filter_table)
        self.tree.header().setSectionsClickable(True)
        self.tree.header().sectionClicked.connect(self.reset_filter)
        
        # Connect Table Click for Report/PDF
        self.table.itemClicked.connect(self.on_table_click)


    def on_table_click(self, item):
        row = item.row()
        col = item.column()
        
        # 0. Check Neutral File Click (Col 6)
        if col == 6:
            val = item.text()
            path_item = self.table.item(row, 11)  # Path column
            if val and path_item:
                folder_path = path_item.text()
                neutral_file_path = os.path.join(folder_path, val)
                if os.path.exists(neutral_file_path):
                    self.cad_viewer.load_neutral_file(neutral_file_path)
                    
                    # Also load CAD-BOM report to mark Only CAD components
                    pcb_item = self.table.item(row, 1)
                    smd_item = self.table.item(row, 3)
                    if pcb_item and smd_item:
                        pcb_code = pcb_item.text().strip()
                        smd_code = smd_item.text().strip()
                        only_cad_refs = self._get_only_cad_refs(smd_code, pcb_code)
                        if only_cad_refs:
                            self.cad_viewer.mark_only_cad_components(only_cad_refs)
                    
                    self.bottom_stack.setCurrentWidget(self.cad_viewer)
                    return
                else:
                    print(f"Neutral file not found: {neutral_file_path}")
        
        # 1. Check WorkSpec Click (Col 10)
        if col == 10:
            val = item.text()
            path_item = self.table.item(row, 11)
            if val and path_item:
                folder_path = path_item.text()
                # Split by newline (previously comma) and filter empties
                pdf_list = [x.strip() for x in val.replace(',', '\n').split('\n') if x.strip()]
                
                selected_pdf = None
                if len(pdf_list) == 1:
                    selected_pdf = pdf_list[0]
                elif len(pdf_list) > 1:
                    # Show Context Menu for Selection
                    menu = QMenu(self)
                    # Add Title or something? No, just list.
                    for p in pdf_list:
                        action = menu.addAction(p)
                        action.setData(p)
                    
                    # Execute Menu at Mouse Position
                    action = menu.exec(QCursor.pos())
                    if action:
                        selected_pdf = action.data()
                
                if selected_pdf:
                    try:
                        folder_path = os.path.abspath(folder_path)
                        pdf_path = os.path.join(folder_path, "WorkSpec", selected_pdf)
                        pdf_path = os.path.normpath(pdf_path)
                        
                        if os.path.exists(pdf_path):
                            print(f"Loading PDF: {pdf_path}")
                            self.web_view.setUrl(QUrl.fromLocalFile(pdf_path))
                            self.bottom_stack.setCurrentWidget(self.web_view)
                            return
                        else:
                            QMessageBox.warning(self, "Error", f"PDF File not found:\n{pdf_path}")
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to load PDF: {e}")
                        print(f"PDF Load Error: {e}")
                        return
        
        # 2. Default: Load Report
        self.bottom_stack.setCurrentWidget(self.report_tabs)
        
        # Cols: 1=PCB, 3=SMD
        pcb_item = self.table.item(row, 1)
        smd_item = self.table.item(row, 3)
        
        if not pcb_item or not smd_item: return
        
        pcb_code = pcb_item.text().strip()
        smd_code = smd_item.text().strip()
        
        report_dir = r"L:\CADBomReport"
        target_prefix = f"{smd_code}_{pcb_code}"
        found_file = None
        
        if os.path.exists(report_dir):
            try:
                for f in os.listdir(report_dir):
                    if f.startswith(target_prefix) and f.lower().endswith(('.xlsx', '.xls')):
                        found_file = os.path.join(report_dir, f)
                        break
            except Exception as e:
                print(f"Directory scan error: {e}")
                
        self.report_tabs.clear()
        
        if found_file:
            try:
                xls = pd.read_excel(found_file, sheet_name=None, header=None)
                for i, (sheet_name, df) in enumerate(xls.items()):
                    # Dynamic Header Parsing Logic (Targeting primarily the first sheet)
                    if i == 0: 
                        header_row_idx = -1
                        # Search first 20 rows
                        for r in range(min(20, len(df))):
                            row_vals = [str(x).strip() for x in df.iloc[r].values]
                            # Check for key columns
                            if "No" in row_vals and "Item" in row_vals and "Result" in row_vals:
                                header_row_idx = r
                                break
                        
                        if header_row_idx != -1:
                            # Set header
                            new_header = df.iloc[header_row_idx]
                            df = df[header_row_idx+1:].copy()
                            df.columns = new_header
                            df.reset_index(drop=True, inplace=True)
                    
                    tab = QWidget()
                    lay = QVBoxLayout(tab)
                    lay.setContentsMargins(2,2,2,2)
                    
                    tv = QTableView()
                    model = PandasModel(df)
                    tv.setModel(model)
                    tv.setAlternatingRowColors(True)
                    tv.horizontalHeader().setStretchLastSection(True)
                    tv.resizeColumnsToContents()
                    
                    lay.addWidget(tv)
                    self.report_tabs.addTab(tab, sheet_name)
            except Exception as e:
                lbl = QLabel(f"Error loading report: {e}")
                self.report_tabs.addTab(lbl, "Error")
        else:
            # Optional: Feedback if needed
            self.report_tabs.addTab(QLabel(f"Report not found for {target_prefix}"), "Info")
    
    def _get_only_cad_refs(self, smd_code, pcb_code):
        """Extract 'Only CAD' component references from CAD-BOM report"""
        only_cad_refs = []
        
        report_dir = r"L:\CADBomReport"
        target_prefix = f"{smd_code}_{pcb_code}"
        found_file = None
        
        print(f"Looking for CAD-BOM report with prefix: {target_prefix}")
        
        if os.path.exists(report_dir):
            try:
                for f in os.listdir(report_dir):
                    if f.startswith(target_prefix) and f.lower().endswith(('.xlsx', '.xls')):
                        found_file = os.path.join(report_dir, f)
                        print(f"Found CAD-BOM report: {found_file}")
                        break
            except Exception as e:
                print(f"Directory scan error: {e}")
                return only_cad_refs
        
        if not found_file:
            print(f"No CAD-BOM report found for {target_prefix}")
            return only_cad_refs
            
        try:
            # Read first sheet with header detection
            xls = pd.read_excel(found_file, sheet_name=0, header=None)
            df = xls
            
            # Find header row
            header_row_idx = -1
            for r in range(min(20, len(df))):
                row_vals = [str(x).strip() for x in df.iloc[r].values]
                if "No" in row_vals and "Item" in row_vals:
                    header_row_idx = r
                    print(f"Found header row at index {r}: {row_vals}")
                    break
            
            if header_row_idx != -1:
                new_header = df.iloc[header_row_idx]
                df = df[header_row_idx+1:].copy()
                df.columns = new_header
                df.reset_index(drop=True, inplace=True)
                
                # Find columns by name (case-insensitive)
                item_col = None
                location_col = None
                for i, col_name in enumerate(df.columns):
                    col_str = str(col_name).strip().lower()
                    if col_str == 'item':
                        item_col = i
                    elif col_str == 'location':
                        location_col = i
                
                print(f"Item column: {item_col}, Location column: {location_col}")
                
                if item_col is not None and location_col is not None:
                    # Extract "Only CAD" rows (case-insensitive)
                    for idx, row in df.iterrows():
                        item_val = str(row.iloc[item_col]).strip().lower()
                        if 'only cad' in item_val:
                            location_val = str(row.iloc[location_col]).strip()
                            if location_val and location_val.lower() != 'nan':
                                # Location is typically a single reference like C1, C2
                                only_cad_refs.append(location_val)
                                print(f"Found Only CAD: {location_val}")
            
            print(f"Total Only CAD components found: {len(only_cad_refs)} - {only_cad_refs}")
        except Exception as e:
            print(f"Error reading CAD-BOM report: {e}")
            import traceback
            traceback.print_exc()
        
        return only_cad_refs


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
            bom_files = [m.get('matrFileNm', '') for m in matr_list if isinstance(m, dict)]
            bom_str = "\n".join(bom_files)
            
            # Check WorkSpec PDF
            work_spec_files = []
            try:
                json_dir = os.path.dirname(file_path)
                ws_dir = os.path.join(json_dir, "WorkSpec")
                if os.path.exists(ws_dir):
                    for f in os.listdir(ws_dir):
                        if f.lower().endswith('.pdf'):
                            work_spec_files.append(f)
            except:
                pass
            
            return {
                'pcbCode': basic.get('pcbCode', ''),
                'rev': basic.get('seq', ''),
                'smdCode': basic.get('smdCode', ''),
                'smdNm': basic.get('smdNm', ''),
                'pcbSize': basic.get('pcbSize', ''),
                'neutralFileNm': cad.get('neutralFileNm', ''),
                'gerberFileNm': cad.get('gerberFileNm', ''),
                'bomFiles': bom_str,
                'matrCount': len(matr_list),
                'workSpecs': "\n".join(work_spec_files)
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
                # Get Path from column 11
                path_item = self.table.item(row, 11)
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
            self.table.setItem(i, 7, QTableWidgetItem(data['gerberFileNm']))
            self.table.setItem(i, 8, QTableWidgetItem(data['bomFiles']))
            self.table.setItem(i, 9, QTableWidgetItem(str(data['matrCount'])))
            self.table.setItem(i, 10, QTableWidgetItem(data.get('workSpecs', '')))
            self.table.setItem(i, 11, QTableWidgetItem(data['FolderPath']))

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
        self.start_date.setDate(QDate(2025, 12, 14))
        h1.addWidget(self.start_date)
        self.layout.addLayout(h1)
        
        # End Date
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("End Date:  "))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setDate(QDate(2025, 12, 15))
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
        self.resize(2048, 1024)
        
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
