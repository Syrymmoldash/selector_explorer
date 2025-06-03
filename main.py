from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QHBoxLayout, QTableWidget, QTableWidgetItem, QComboBox,
    QCheckBox, QHeaderView, QLabel, QTextEdit, QAbstractItemView, QPushButton, QMessageBox, QSizePolicy
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QFontMetrics, QPainter, QPen, QIcon
import time
from pywinauto import uia_element_info
from pywinauto.application import WindowSpecification
import json
import threading
import os
import win32process
import win32gui
import pywinauto
import pyautogui
import keyboard
import sys


CONTROL_TYPE_MAPPING = {
    'Dialog': 'Window',
    'Dlg': 'Window',
    'Textbox': 'Edit',
    'Listview': 'List',
    'Listitem': 'ListItem',
    'Radio': 'RadioButton',
    'Dropdown': 'ComboBox',
    'Treeview': 'Tree',
    'Label': 'Text',
    'Panel': 'Pane'
}


def main_process():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("pythonrpa_logo.ico")))
    window = SelectorExplorer()
    window.show()
    app.exec()

def resource_path(relative_path):
    """ Get absolute path to resource, works for PyInstaller bundle """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return relative_path


def find_element(selector):
    results = find_elements(selector)
    if results:
        return results[0]
    return False


def find_matches(parent, selector, level=0):
    if level >= len(selector):
        return [parent]
    matches = []
    expected_props = selector[level]
    if not hasattr(parent, "children"):
        return []
    try:
        children = parent.children()
    except AttributeError:
        return []

    def selector_params(params: dict, child):
        props = {
            "title": child.window_text(),
            "class_name": child.class_name(),
            "control_type": child.friendly_class_name(),
            "control_id": child.control_id(),
            "ctrl_index": child.parent().children().index(child)
        }
        if CONTROL_TYPE_MAPPING.get(props["control_type"]):
            props["control_type"] = CONTROL_TYPE_MAPPING.get(props["control_type"])
        for v in params:
            if params[v] != props[v]:
                return False
        return True


    for child in children:
        element_found = selector_params(expected_props, child)
        if element_found:
            matched_descendants = find_matches(child, selector, level + 1)
            matches.extend(matched_descendants)
    return matches


def find_elements(selector):
    if not selector or not isinstance(selector[0], dict):
        return []
    window = pywinauto.Desktop("uia").window(**selector[0])
    if not window.exists():
        return []
    matching_elements = []

    try:
        matching_elements.extend(find_matches(window, selector[1:]))
    except Exception as e:
        return []
    return matching_elements


def safe_compare(element1, element2):
    if element1 is None or element2 is None:
        return False  # Avoid comparing None elements
    try:
        return bool(uia_element_info.IUIA().iuia.CompareElements(element1.element, element2.element))
    except Exception as e:
        return False  # Handle invalid elements safely


def wait_element_to_appear(selector: list, wait_time: int=30):
    start_time = time.time()
    result = None
    while True:
        result = find_elements(selector)
        if result:
            return True
        if time.time() - start_time > wait_time:
            break
    if result:
        return True
    else:
        return False


def wait_element_to_disappear(selector: list, wait_time: int=30):
    start_time = time.time()
    result = None
    while True:
        result = find_elements(selector)
        if not result:
            return True
        if time.time() - start_time > wait_time:
            break
    if not result:
        return True
    else:
        return False


class HighlightWindow(QWidget):
    finished = pyqtSignal(list)
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint |
                            Qt.WindowType.Tool | Qt.WindowType.BypassWindowManagerHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)  # Transparent background
        self.highlight_pid = self.get_window_pid()
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_highlight)
        self.timer.start(100)  # Slightly slower refresh to prevent overload

        self.rect_x, self.rect_y, self.rect_w, self.rect_h = 0, 0, 0, 0  # Default rectangle
        self.last_element = None  # Store last highlighted element
        self.element = None
        self.deepest_element = None

    def get_window_pid(self):
        """Finds and stores the PID of the highlight window."""
        hwnd = self.winId()  # Get the window handle
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid  # Return the PID for filtering

    def get_foreground_window(self):
        hwnd = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(hwnd)

    def get_element(self, x, y):
        """Finds the first valid UI element that is NOT from the ignored process."""
        element = pywinauto.Desktop(backend="uia").from_point(x, y)
        while element and element.element_info.process_id == self.highlight_pid:
            element = element.parent()  # Move up the UI tree to find a valid element
        foreground_title = self.get_foreground_window()
        if element.top_level_parent().element_info.name != foreground_title:
            return None
        return element

    def get_selector(self, element):
        selector = []
        while element != element.top_level_parent():
            props = {
                "title": element.window_text(),
                "class_name": element.class_name(),
                "control_type": element.friendly_class_name(),
                "control_id": element.control_id(),
                "ctrl_index": element.parent().children().index(element),
                "iface": self.get_elements_functions(element)
            }
            selector.append(props)
            element = element.parent()
        props = {
            "title": element.window_text(),
            "class_name": element.class_name(),
            "control_type": element.friendly_class_name(),
            "control_id": element.control_id(),
            "iface": self.get_elements_functions(element)
        }
        selector.append(props)
        level = 0
        selector = selector[::-1]
        for elm_dict in selector:
            if CONTROL_TYPE_MAPPING.get(elm_dict["control_type"]):
                elm_dict["control_type"] = CONTROL_TYPE_MAPPING[elm_dict["control_type"]]
            elm_dict["level"] = level
            level = level + 1
        return selector

    def get_elements_functions(self, element):
        element_funcs = {}
        iface_mapping = {
            "iface_value": lambda el: f"Current Value: {el.iface_value.CurrentValue}, Current Is Read Only: {el.iface_value.CurrentIsReadOnly}",
            "iface_selection": lambda el: f"Current Selection: {[el.iface_selection.GetCurrentSelection().GetElement(i).CurrentName for i in range(el.iface_selection.GetCurrentSelection().Length)]}, Selection Is Required: {el.iface_selection.CurrentIsSelectionRequired}, Current Can Select Multiply: {el.iface_selection.CurrentCanSelectMultiple}",
            "iface_toggle": lambda el: f"Toggle State: {el.iface_toggle.CurrentToggleState}",
            "iface_scroll": lambda el: f"Current Horizontal Scroll Percent: {el.iface_scroll.CurrentHorizontalScrollPercent}, Current Vertical Scroll Percent: {el.iface_scroll.CurrentVerticalScrollPercent}, Current Horizontally Scrollable: {el.iface_scroll.CurrentHorizontallyScrollable}, Current Vertically Scrollable: {el.iface_scroll.CurrentVerticallyScrollable}, Current Horizontal View Size: {el.iface_scroll.CurrentHorizontalViewSize}, Current Vertical View Size: {el.iface_scroll.CurrentVerticalViewSize}",
            "iface_invoke": lambda el: f"Clickable: Yes",
            "iface_range_value": lambda el: f"Current Value: {el.iface_range_value.CurrentValue}, Current Is Read Only: {el.iface_range_value.CurrentIsReadOnly}, Current Maximum: {el.iface_range_value.CurrentMaximum}, Current Minimum: {el.iface_range_value.CurrentMinimum}, Current Large Change: {el.iface_range_value.CurrentLargeChange}, Current Small Change: {el.iface_range_value.CurrentSmallChange}",
            "iface_text": lambda el: f"Current Text: {el.DocumentRange.GetText(-1)}",
            "iface_grid": lambda el: f"Current Row Count: {el.iface_grid.CurrentRowCount}, Current Column Count: {el.iface_grid.CurrentColumnCount}",  # Grid dimensions
            "iface_table": lambda el: f"Current Row Count: {el.iface_table.CurrentRowCount}, Current Column Count: {el.iface_table.CurrentColumnCount}, Current Row Or Column Major: {el.iface_table.CurrentRowOrColumnMajor}, Current Row Headers: {[el.iface_table.GetCurrentRowHeaders().GetElement(i).CurrentName for i in range(el.iface_table.GetCurrentRowHeaders().Length)]}, Current Column Headers: {[el.iface_table.GetCurrentColumnHeaders().GetElement(i).CurrentName for i in range(el.iface_table.GetCurrentColumnHeaders().Length)]}",  # Table headers
            "iface_expand_collapse": lambda el: f"Expand Status: {el.iface_expand_collapse.CurrentExpandCollapseState}",
            "iface_window": lambda el: f"Can be maximized: {el.iface_window.CurrentCanMaximize}, Can be minimized: {el.iface_window.CurrentCanMinimize}, Is Modal: {el.iface_window.CurrentIsModal}, Current Is Topmost: {el.iface_window.CurrentIsTopmost}, Current Window Visual State: {el.iface_window.CurrentWindowVisualState}, Current Window Interaction State: {el.iface_window.CurrentWindowInteractionState}",
            "iface_transform": lambda el: f"Can Move: {el.iface_transform.CurrentCanMove}, Can Resize: {el.iface_transform.CurrentCanResize}, Can Rotate: {el.iface_transform.CurrentCanRotate}",  # Element bounding box
            "iface_drag": lambda el: f"Is Grabbed: {el.iface_drag.CurrentIsGrabbed}, Drop Effect: {el.iface_drag.CurrentDropEffect}, Drop Effects: {el.iface_drag.CurrentDropEffects}",  # Dragging support
            "iface_drop_target": lambda el: f"Drop Target Effect: {el.iface_drop_target.CurrentDropTargetEffect}, Drop Target Effect: {el.iface_drop_target.CurrentDropTargetEffects}",  # Dropping support
        }
        for iface, func in iface_mapping.items():
            try:
                if hasattr(element, iface):
                    element_funcs[iface.replace("iface_", "").capitalize()] = func(element)
            except:
                pass
        return element_funcs

    def get_deepest_element(self, element, x, y, max_depth=10):
        """Finds the smallest child element under the cursor, preventing infinite recursion."""
        stack = [(element, 0)]  # Use stack with depth tracking
        last_valid_element = element  # Store last valid non-empty element

        while stack:
            current_element, depth = stack.pop()
            if depth >= max_depth:  # Prevent infinite recursion
                break

            children = current_element.children()

            # **Skip elements from the highlight window**
            if current_element.element_info.process_id == self.highlight_pid:
                continue  # Ignore this element and move to the next one

            # **If no children, return the current valid element**
            if not children:
                return current_element

            # **Check deeper elements**
            for child in children:
                rect = child.element_info.rectangle
                if rect.left <= x <= rect.right and rect.top <= y <= rect.bottom:
                    stack.append((child, depth + 1))
                    last_valid_element = child  # Store last valid element

        return last_valid_element  # Return last valid element if no deeper one found

    def update_highlight(self):
        global selector
        selector.clear()
        if keyboard.is_pressed("esc"):
            self.finished.emit([])
            self.close()
            return

        """Detects the topmost hovered element and updates the highlight position."""
        x, y = pyautogui.position()
        try:
            self.element = self.get_element(x, y)
        except Exception as e:
            return

        if not self.element:
            self.hide()
            return

        # **Find the absolute deepest child under the cursor**
        self.deepest_element = self.get_deepest_element(self.element, x, y)

        # **Check if Ctrl is pressed and return ancestor properties**
        if keyboard.is_pressed("ctrl"):
            selector = self.get_selector(self.deepest_element)
            self.finished.emit(selector)
            self.timer.stop()
            self.hide()

        rect = self.deepest_element.element_info.rectangle
        # **Force update even if the new element is inside the previous one**
        if rect and rect.right > rect.left and rect.bottom > rect.top:
            new_rect = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)

            # **Prevent highlighting oversized elements (e.g., desktop)**
            if new_rect[2] > screen_width or new_rect[3] > screen_height:
                self.hide()
                return

            # **Ensure update if switching elements inside the same area**
            if not safe_compare(self.deepest_element, self.last_element):
                self.rect_x, self.rect_y, self.rect_w, self.rect_h = new_rect
                self.setGeometry(self.rect_x, self.rect_y, self.rect_w, self.rect_h)
                self.last_element = self.deepest_element  # Save the last element
                self.show()
                self.update()  # Force repaint

        QApplication.processEvents()  # Prevent UI freeze

    def paintEvent(self, event):
        """Draws the green border around the hovered element."""
        painter = QPainter(self)
        pen = QPen(Qt.GlobalColor.green, 3)  # Green border, 3px width
        painter.setPen(pen)
        painter.drawRect(0, 0, self.rect_w - 1, self.rect_h - 1)  # Draw border


class SelectorExplorer(QMainWindow):
    global selector
    open_highlight_signal = pyqtSignal()
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Selector Explorer")
        self.setGeometry(100, 100, 1000, 600)
        self.setWindowIcon(QIcon(resource_path("pythonrpa_logo.ico")))
        self.highlight_window = None
        self.open_highlight_signal.connect(self.open_highlight_window)

        self.selected_element = None
        self.checkboxes = {}
        self.current_actions = {}
        self.result_return_actions_list = ["GetCellValue(row, column)"]
        self.actions_dict = {
            "Click": lambda arg: arg[0].click_input(),
            "DoubleClick": lambda arg: arg[0].double_click_input(),
            "Value": {
                "SetValue(value)": lambda arg: arg[0].iface_value.SetValue(arg[1])
            },
            "Selection": {
                "Select": lambda arg: arg[0].iface_selection_item.Select()
            },
            "Toggle": {
                "Toggle": lambda arg: arg[0].iface_toggle.Toggle()
            },
            "Scroll": {
                "ScrollByPercent(x, y)": lambda arg: arg[0].iface_scroll.SetScrollPercent(float(arg[1]), float(arg[2])),
                "ScrollStep(x, y)": lambda arg: arg[0].iface_scroll.Scroll(int(arg[1]), int(arg[2]))
            },
            "Range_value": {
                "SetValue(Value)": lambda arg: arg[0].iface_range_value.SetValue(float(arg[1]))
            },
            "Grid": {
                "GetCellValue(row, column)": lambda arg: arg[0].iface_grid.GetItem(int(arg[1]), int(arg[2])).CurrentName
            },
            "Expand_collapse": {
                "Expand": lambda arg: arg[0].iface_expand_collapse.Expand(),
                "Collapse": lambda arg: arg[0].iface_expand_collapse.Collapse(),
            },
            "Window": {
                "Close": lambda arg: arg[0].iface_window.Close(),
                "WindowState(0-Normal, 1-Maximize, 2-Minimize)": lambda arg: arg[0].iface_window.SetWindowVisualState(int(arg[1]))
            },
            "Transform": {
                "Move(x, y)": lambda arg: arg[0].iface_transform.Move(int(arg[1]), int(arg[2])),
                "Resize(width, height)": lambda arg: arg[0].iface_transform.Resize(int(arg[1]), int(arg[2]))
            }
        }

        self.setStyleSheet("""
            /* General Styles */
            QWidget {
                background-color: #F9F9F9;  /* Clean white background */
                color: #1A2038;  /* Dark blue-gray text */
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 14px;
            }

            /* Tree Widget, Table Widget, and Text Edit */
            QTreeWidget, QTableWidget, QTextEdit {
                background-color: #FFFFFF;  /* Pure white panels */
                border: 1px solid #E0E0E0;  /* Soft gray borders */
                border-radius: 8px;
                padding: 6px;
            }

            /* Buttons */
            QPushButton {
                background-color: #1A2038;  /* Deep navy blue */
                color: white;
                border-radius: 8px;
                padding: 12px;
                font-size: 14px;
                font-weight: bold;
                border: none;
            }

            QPushButton:hover {
                background-color: #232A48;  /* Slightly darker navy on hover */
            }

            QPushButton:pressed {
                background-color: #0F1428;  /* Even darker on press */
            }

            /* Headers */
            QHeaderView::section {
                background-color: #E0E0E0;  /* Light gray headers */
                color: #1A2038;
                padding: 8px;
                font-weight: bold;
                border-radius: 6px;
            }

            /* Checkboxes */
            QCheckBox {
                spacing: 8px;
            }

            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                background-color: #E0E0E0;  /* Light gray */
                border: 2px solid #1A2038;
            }

            QCheckBox::indicator:checked {
                background-color: #1A2038;  /* Dark blue */
                border: 2px solid white;
            }

            /* Scrollbars */
            QScrollBar:vertical {
                background: #F0F0F0;
                width: 12px;
                margin: 2px 0 2px 0;
                border-radius: 6px;
            }

            QScrollBar::handle:vertical {
                background: #C0C0C0;
                border-radius: 6px;
            }

            QScrollBar::handle:vertical:hover {
                background: #A0A0A0;
            }

            QScrollBar::sub-line:vertical, QScrollBar::add-line:vertical {
                background: none;
                border: none;
            }
        """)

        # Layouts
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        bottom_layout = QHBoxLayout()
        left_bottom_layout = QVBoxLayout()
        left_bottom_labels_layout = QHBoxLayout()
        right_bottom_layout = QVBoxLayout()
        bottom_layout.addLayout(left_bottom_layout)
        bottom_layout.addLayout(right_bottom_layout)
        bottom_layout.setStretchFactor(left_bottom_layout, 3)
        bottom_layout.setStretchFactor(right_bottom_layout, 2)
        main_layout.addLayout(top_layout, 5)
        main_layout.addLayout(bottom_layout, 4)

        # Left Panel (Tree View) with Scrollbars
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Elements Tree")
        self.tree.itemClicked.connect(self.on_element_selected)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.tree.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)  # Smooth scrolling
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # Allow user resizing
        self.tree.setIndentation(10)  # Proper indentation for child elements
        top_layout.addWidget(self.tree)

        # Mid Panel (Search Properties) with Scrollbars
        self.properties_table = QTableWidget()
        self.properties_table.setColumnCount(2)
        self.properties_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.properties_table.setHorizontalHeaderLabels(["", "Search Properties"])
        self.properties_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.properties_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.properties_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.properties_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.properties_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        top_layout.addWidget(self.properties_table)

        # Right Panel (Legacy Properties) with Scrollbars
        self.full_props_table = QTableWidget()
        self.full_props_table.setColumnCount(1)
        self.full_props_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.full_props_table.setHorizontalHeaderLabels(["Full Properties"])
        self.full_props_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.full_props_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.full_props_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.full_props_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        top_layout.addWidget(self.full_props_table)

        # Bottom Panel (Selector Display) with Scrollbars
        self.selector_display = QTextEdit()
        self.selector_display.setReadOnly(True)
        self.selector_display.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.selector_display.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.selector_display.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        left_bottom_labels_layout.addWidget(QLabel("Generated Selector:"), 1)
        # ✅ Button
        self.copy_button = QPushButton("Copy")
        self.copy_button.clicked.connect(self.copy_selector_to_clipboard)
        self.copy_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        left_bottom_labels_layout.addWidget(self.copy_button, 1)
        left_bottom_layout.addLayout(left_bottom_labels_layout, 1)
        left_bottom_layout.addWidget(self.selector_display, 8)

        self.actions_label = QLabel("Actions list")
        right_bottom_layout.addWidget(self.actions_label)
        self.dropdown = QComboBox()
        self.dropdown.addItems([])
        self.dropdown.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_bottom_layout.addWidget(self.dropdown, 1)  # ✅ Dropdown at top-right of bottom panel

        self.arg_label = QLabel("Enter arguments(Separate by comma)")
        right_bottom_layout.addWidget(self.arg_label)

        self.arg_input = QTextEdit()
        right_bottom_layout.addWidget(self.arg_input, 2)

        # ✅ Button
        self.action_button = QPushButton("Perform Action")
        self.action_button.clicked.connect(self.on_perform)
        self.action_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_bottom_layout.addWidget(self.action_button, 2)  # ✅ Button below dropdown

        # Cancel Button
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.on_cancel)
        self.cancel_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_bottom_layout.addWidget(self.cancel_button, 2)

        # Explore Button
        self.explore_button = QPushButton("Explore")
        self.explore_button.clicked.connect(self.on_explorer)
        self.explore_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_bottom_layout.addWidget(self.explore_button, 2)

        self.properties_table.clearContents()
        self.properties_table.setRowCount(0)
        self.full_props_table.clearContents()
        self.full_props_table.setRowCount(0)

        # Populate tree
        self.populate_tree()

    def populate_tree(self):
        global selector
        self.tree.clear()
        item_map = {}
        max_width = 0

        for level, element in enumerate(selector):
            tree_text = (
                f"Level {element.get('level')}, Title: {element.get('title')}, "
                f"class_name: {element.get('class_name')}"
            )
            font_metrics = QFontMetrics(self.tree.font())
            text_width = font_metrics.horizontalAdvance(tree_text) + font_metrics.horizontalAdvance(" ") * 25
            max_width = max(max_width, text_width + level * 10)

            # Create a tree item
            item = QTreeWidgetItem([tree_text])
            item.setData(0, 1, element)  # Store element data for selection

            # Check if this element has a parent (previous level element)
            if level > 0 and (level - 1) in item_map:
                # Add this as a child to its parent
                parent_item = item_map[level - 1]
                parent_item.addChild(item)
            else:
                # This is a top-level item
                self.tree.addTopLevelItem(item)

            # Store the item in the map
            item_map[level] = item
            self.checkboxes[level] = ["title", "class_name", "control_type"]

        # Expand all nodes initially
        self.tree.expandAll()
        self.adjust_column_width("tree", max_width)

    def adjust_column_width(self, widget_type: str, max_width: int):
        if widget_type == "tree":
            self.tree.setColumnWidth(0, max_width)
        elif widget_type == "table1":
            self.properties_table.setColumnWidth(1, max_width)
        elif widget_type == "table2":
            self.full_props_table.setColumnWidth(0, max_width)

    def on_element_selected(self, item):
        self.selected_element = item.data(0, 1)
        if self.selected_element:
            self.update_properties_table()
            self.update_full_props_table()

    def create_checkbox_callback(self, key):
        """Returns a proper callback function to prevent lambda closure issues."""
        return lambda state, key_copy=key: self.on_checkbox_state_changed(key_copy, state)

    def update_properties_table(self):
        """Updates the properties table when an element is selected."""
        if not self.selected_element:
            return

        self.properties_table.clearContents()

        levels_filter = [key for key in self.selected_element if key not in ["level", "iface"]]
        self.properties_table.setRowCount(len(levels_filter))
        row = 0
        width_size = 0
        for key, value in self.selected_element.items():
            if key not in levels_filter:
                continue
            checkbox = QCheckBox()
            level = self.selected_element.get("level")
            font_metrics = QFontMetrics(self.properties_table.font())
            text_width = font_metrics.horizontalAdvance(key + str(value)) + 10
            width_size = max(width_size, text_width)
            if level is not None:
                # Check if the key is in the checkboxes dictionary for this level
                checked = key in self.checkboxes.get(level, [])
                checkbox.setChecked(checked)
                # Connect the checkbox state change to the callback
                checkbox.stateChanged.connect(self.create_checkbox_callback(key))

            self.properties_table.setCellWidget(row, 0, checkbox)
            self.properties_table.setItem(row, 1, QTableWidgetItem(f"{key}: {value}"))
            row += 1
        self.adjust_column_width("table1", width_size)
        self.update_selector()

    def update_full_props_table(self):
        """Updates the properties table when an element is selected."""
        if not self.selected_element:
            return

        self.full_props_table.clearContents()
        row_counts = len(self.selected_element.get("iface"))
        self.full_props_table.setRowCount(row_counts)
        row = 0
        width_size = 0
        full_property_dict = self.selected_element.get("iface").copy()
        for key, value in full_property_dict.items():
            font_metrics = QFontMetrics(self.full_props_table.font())
            text_width = font_metrics.horizontalAdvance(key + str(value)) + 10
            width_size = max(width_size, text_width)
            self.full_props_table.setItem(row, 0, QTableWidgetItem(f"{key}: {value}"))
            row += 1
        self.adjust_column_width("table2", width_size)
        self.update_actions_list(full_property_dict)
        self.update_selector()

    def update_actions_list(self, full_property_dict):
        self.dropdown.clear()
        for key, value in self.actions_dict.items():
            if full_property_dict.get(key):
                for action_type, action_func in self.actions_dict[key].items():
                    self.current_actions[action_type] = action_func
                    self.dropdown.addItem(action_type)
            else:
                self.current_actions[key] = value
        self.dropdown.addItem("Click")
        self.dropdown.addItem("DoubleClick")

    def on_checkbox_state_changed(self, key, state):
        """Handles checkbox state changes and updates the checkboxes dictionary."""
        if not self.selected_element:
            return

        level = self.selected_element.get("level")

        # Get the current keys for this level or an empty list if not found
        current_keys = self.checkboxes.get(level, [])

        # Update the keys based on the checkbox state
        if state == Qt.CheckState.Checked.value:
            if key not in current_keys:
                current_keys.append(key)
        else:
            if key in current_keys:
                current_keys.remove(key)

        # Update the checkboxes dictionary
        self.checkboxes[level] = current_keys
        self.update_selector()

    def update_selector(self):
        """Updates the bottom panel with the current selector."""
        global selector, find_selector
        if not self.selected_element:
            return

        # Find the index of the selected element in the global selector
        index = selector.index(self.selected_element)
        selected_subset = selector[: index + 1]
        # Filter the elements based on the checkboxes dictionary
        selector_text = [
            {k: v for k, v in element.items() if k in self.checkboxes.get(element.get("level"), [])}
            for element in selected_subset
        ]

        # Update the global selector
        find_selector = selector_text.copy()

        # Update the bottom panel
        self.selector_display.clear()
        self.selector_display.setText(json.dumps(find_selector, ensure_ascii=False, indent=4))

    def on_cancel(self):
        global selector
        selector.clear()
        QApplication.quit()

    def on_explorer(self):
        self.hide()
        self.open_highlight_signal.emit()

    def on_perform(self):
        global find_selector
        selector_field = self.selector_display.toPlainText().strip()
        if not selector_field:
            QMessageBox.warning(self, "Warning", "Selector is not chosen")
            return
        action = self.dropdown.currentText()
        if not action:
            QMessageBox.warning(self, "Warning", "Action is not chosen")
            return
        element = find_element(find_selector)
        if type(element) == WindowSpecification:
            element = element.wrapper_object()
        if not element:
            QMessageBox.warning(self, "Warning", "Element is not found")
            return
        else:
            args = [element] + [arg for arg in str(self.arg_input.toPlainText().strip()).split(",") if len(arg) != 0]
            self.hide()
            try:
                if action in self.result_return_actions_list:
                    QMessageBox.warning(self, "Result", str(self.current_actions[action](args)))
                else:
                    self.current_actions[action](args)
            except Exception as e:
                QMessageBox.warning(self, "Warning", "Action failed")
            self.show()

    def open_highlight_window(self):
        if not self.highlight_window or self.highlight_window.isHidden():
            self.highlight_window = HighlightWindow()
            self.highlight_window.finished.connect(self.on_highlight_finished)
        self.highlight_window.show()

    def copy_selector_to_clipboard(self):
        text = self.selector_display.toPlainText().strip()
        if text:
            QTimer.singleShot(0, lambda: QApplication.clipboard().setText(text))

    def on_highlight_finished(self, selector_elements):
        global selector
        if selector_elements:
            selector = selector_elements
        self.showNormal()
        self.properties_table.clearContents()
        self.properties_table.setRowCount(0)
        self.populate_tree()
        self.highlight_window = None


if __name__ == "__main__":
    selector = []
    find_selector = []
    screen_width, screen_height = pyautogui.size()
    qt_thread = threading.Thread(target=main_process, daemon=True)
    qt_thread.start()
    qt_thread.join()
