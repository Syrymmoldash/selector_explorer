from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QHBoxLayout, QTableWidget, QTableWidgetItem, QComboBox,
    QCheckBox, QHeaderView, QLabel, QTextEdit, QAbstractItemView, QPushButton, QMessageBox, QSizePolicy, QLineEdit
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QFontMetrics, QPainter, QPen, QIcon, QFont
import sys
import time
from pywinauto import uia_element_info
from pywinauto.uia_element_info import UIAElementInfo
from pywinauto.controls.uiawrapper import UIAWrapper
import json
import os
import win32process
import win32gui
import re
import pyautogui
import keyboard
import psutil


def resource_path(relative_path):
    """ Get absolute path to resource, works for PyInstaller bundle """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return relative_path


def find_element(selector):
    result = find_elements(selector)
    if result:
        return UIAWrapper(result[0])
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
    if expected_props.get("ctrl_index"):
        try:
            matched_descendants = find_matches(children[expected_props.get("ctrl_index")], selector, level + 1)
            matches.extend(matched_descendants)
            return matches
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            return []

    def selector_params(params: dict, child):
        properties = {}
        try:
            properties = child.get_properties()
            properties["title"] = child.element_info.name
            properties.pop("rectangle")
        except Exception as e:
            properties = {
                "class_name": child.element_info.class_name,
                "title": child.element_info.name,
                "control_type": child.element_info.control_type,
                "rich_text": child.element_info.rich_text,
                "visible": child.element_info.visible,
                "enabled": child.element_info.enabled,
                "control_id": child.element_info.control_id,
                "automation_id": child.element_info.automation_id
            }
        for v in params:
            if not properties.get(v):
                continue
            if params[v] != properties[v]:
                return False
        return True


    for child in children:
        element_found = selector_params(expected_props, UIAWrapper(child))
        if element_found:
            matched_descendants = find_matches(child, selector, level + 1)
            matches.extend(matched_descendants)
    return matches


def find_elements(selector):
    if not selector or not isinstance(selector[0], dict):
        return []
    desktop = UIAElementInfo()
    window_search_args_list = ["process", "class_name", "title", "control_type", "content_only", "title_re"]
    window_search_args_dict = {}
    for k, v in selector[0].items():
        if k in window_search_args_list:
            window_search_args_dict[k] = v
    if window_search_args_dict.get("title_re"):
        matching_windows = [child for child in desktop.children() if re.match(window_search_args_dict["title_re"], child.name)]
        if matching_windows:
            window_search_args_dict["title"] = matching_windows[0]
            window_search_args_dict.pop("title_re")
    window = desktop.children(**window_search_args_dict)
    if not window:
        return []
    matching_elements = []
    try:
        matching_elements.extend(find_matches(window[0], selector[1:]))
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

        self.highlight_pid = self.get_window_pid()  # Get PID once at startup
        self.backend = "uia"  # Set the backend used

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_highlight)
        self.timer.start(100)  # Slightly slower refresh to prevent overload

        self.rect_x, self.rect_y, self.rect_w, self.rect_h = 0, 0, 0, 0  # Default rectangle
        self.last_element = None  # Store last highlighted element

    def get_process_name(self, process_id):
        """Returns the process name given its ID."""
        try:
            return psutil.Process(process_id).name()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            return None  # Return None if process is unavailable

    def get_foreground_window(self):
        hwnd = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(hwnd)

    def get_window_pid(self):
        """Finds and stores the PID of the highlight window."""
        hwnd = self.winId()  # Get the window handle
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid  # Return the PID for filtering

    def get_application_window(self, element):
        """Finds the top-most window belonging to the same process."""
        while element.parent and element.parent.process_id == element.process_id:
            element = element.parent  # Move up to the main application window
        return element  # Return the application-level window

    def get_filtered_element(self, x, y):
        """Finds the first valid UI element that is NOT from the ignored process."""
        element = UIAElementInfo.from_point(x, y)

        # **Filter out elements from the highlight window itself**
        while element and element.process_id == self.highlight_pid:
            element = element.parent  # Move up the UI tree to find a valid element
        foreground_title = self.get_foreground_window()
        top_window = self.get_application_window(element)
        if top_window.name != foreground_title:
            return None
        return element

    def get_ctrl_index(self, element):
        """Finds the index of an element among all its siblings."""
        parent = element.parent  # Get the parent element
        if not parent:
            return None  # No parent means it's a top-level window

        siblings = parent.children()  # Get all children (siblings)

        try:
            return siblings.index(element)
        except ValueError:
            return None  # If element is not found for some reason


    def get_element_properties(self, element):
        """Extracts all available properties of a UIA element into a dictionary."""
        properties = {}
        try:
            properties = UIAWrapper(element).get_properties()
            properties["title"] = element.name
            properties.pop("rectangle")
        except Exception as e:
            properties = {
                "class_name": element.class_name,
                "title": element.name,
                "control_type": element.control_type,
                "rich_text": element.rich_text,
                "visible": element.visible,
                "enabled": element.enabled,
                "control_id": element.control_id,
                "automation_id": element.automation_id,
            }
        properties["iface"] = {}
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
                if hasattr(UIAWrapper(element), iface):
                    properties["iface"][iface.replace("iface_", "").capitalize()] = func(UIAWrapper(element))
            except:
                pass
        return properties


    def get_ancestor_properties(self, element):
        """Collects properties of the element and all its parents up to the application level."""
        ancestors = []
        application_element = self.get_application_window(element)  # Find the application-level element

        while element and element != application_element:
            props = self.get_element_properties(element)

            # **For all child elements, add ctrl_index**
            ctrl_index = self.get_ctrl_index(element)
            if ctrl_index is not None:
                props["ctrl_index"] = ctrl_index

            ancestors.append(props)
            element = element.parent  # Move to the parent

        # **Set backend only on the application level (top-most application element)**
        app_props = self.get_element_properties(application_element)
        app_props["backend"] = self.backend
        ancestors.append(app_props)  # Add application level as the last element

        return ancestors[::-1]  # Reverse order (application-level first)

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
            if current_element.process_id == self.highlight_pid:
                continue  # Ignore this element and move to the next one

            # **If no children, return the current valid element**
            if not children:
                return current_element

            # **Check deeper elements**
            for child in children:
                rect = child.rectangle
                if rect.left <= x <= rect.right and rect.top <= y <= rect.bottom:
                    stack.append((child, depth + 1))
                    last_valid_element = child  # Store last valid element

        return last_valid_element  # Return last valid element if no deeper one found

    def update_highlight(self):
        global selector
        if keyboard.is_pressed("esc"):
            self.finished.emit([])
            self.close()
            return

        """Detects the topmost hovered element and updates the highlight position."""
        x, y = pyautogui.position()
        try:
            element = self.get_filtered_element(x, y)  # Get element while ignoring the highlight window
        except Exception as e:
            return

        if not element:
            self.hide()
            return

        # **Find the absolute deepest child under the cursor**
        deepest_element = self.get_deepest_element(element, x, y)
        rect = deepest_element.rectangle

        # **Check if Ctrl is pressed and return ancestor properties**
        if keyboard.is_pressed("ctrl"):
            selector.clear()
            selector.extend(self.get_ancestor_properties(deepest_element))
            self.finished.emit(selector)
            self.close()

        # **Force update even if the new element is inside the previous one**
        if rect and rect.right > rect.left and rect.bottom > rect.top:
            new_rect = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)

            # **Prevent highlighting oversized elements (e.g., desktop)**
            if new_rect[2] > screen_width or new_rect[3] > screen_height:
                self.hide()
                return

            # **Ensure update if switching elements inside the same area**
            if not safe_compare(deepest_element, self.last_element):
                self.rect_x, self.rect_y, self.rect_w, self.rect_h = new_rect
                self.setGeometry(self.rect_x, self.rect_y, self.rect_w, self.rect_h)
                self.last_element = deepest_element  # Save the last element
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
    open_highlight_signal = pyqtSignal()
    def __init__(self):
        super().__init__()
        global selector
        self.setWindowTitle("Selector Explorer")
        self.setGeometry(100, 100, 1000, 600)
        self.setWindowIcon(QIcon(resource_path("pythonrpa_logo.ico")))
        self.highlight_window = None
        self.open_highlight_signal.connect(self.open_highlight_window)

        self.selected_element = None
        self.checkboxes = {}
        self.selected_elements_data = []
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
        self.current_actions = {}

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
        left_bottom_layout.addWidget(QLabel("Generated Selector:"))
        left_bottom_layout.addWidget(self.selector_display)

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
        self.tree.clear()
        global selector
        item_map = {}  # Stores elements by their index for parent-child linking
        max_width = 0

        for level, element in enumerate(selector):
            title = element.get('title', 'Unnamed')
            class_name = element.get('class_name', '')
            ctrl_index = element.get('ctrl_index', None)
            backend = element.get('backend', '')
            element["level"] = level

            # Format tree display
            tree_text = f"Level {level}, Title: {title}, class_name: {class_name}"
            if backend:
                tree_text += f", backend: {backend}"
            if ctrl_index is not None:
                tree_text += f", ctrl_index: {ctrl_index}"

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
            if level == 0:
                self.checkboxes[level] = ["title", "backend", "class_name"]
            else:
                self.checkboxes[level] = ["title", "ctrl_index", "class_name"]

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

    def on_element_selected(self, item, column):
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

        levels_filter = []
        if self.selected_element.get("level") == 0:
            levels_filter = [key for key in self.selected_element if key in ["process", "class_name", "title", "control_type", "content_only"]]
        else:
            levels_filter = [key for key in self.selected_element if key in ["class_name", "title", "control_type", "rich_text", "visible", "enabled", "control_id", "automation_id"]]
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
        for key, value in self.selected_element.items():
            if key not in ["class_name", "title", "control_type", "rich_text", "visible", "enabled", "control_id", "automation_id", "process", "content_only"]:
                full_property_dict[key] = value
        for key, value in full_property_dict.items():
            font_metrics = QFontMetrics(self.properties_table.font())
            text_width = font_metrics.horizontalAdvance(key + str(value)) + 10
            width_size = max(width_size, text_width)
            self.full_props_table.setItem(row, 0, QTableWidgetItem(f"{key}: {value}"))
            row += 1
        self.adjust_column_width("table2", width_size)
        self.update_actions_list()
        self.update_selector()

    def update_actions_list(self):
        self.dropdown.clear()
        for key, value in self.actions_dict.items():
            if self.selected_element["iface"].get(key):
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
        if level is None:
            return

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
        global selector, final_selector
        if not self.selected_element:
            return

        try:
            # Find the index of the selected element in the global selector
            index = selector.index(self.selected_element)
            selected_subset = selector[: index + 1]

            # Filter the elements based on the checkboxes dictionary
            selector_text = [
                {k: v for k, v in element.items() if k in self.checkboxes.get(element.get("level"), [])}
                for element in selected_subset
            ]

            # Update the global selector
            final_selector = selector_text.copy()

            # Update the bottom panel
            self.selector_display.clear()
            self.selector_display.setText(json.dumps(final_selector, ensure_ascii=False, indent=4))
        except ValueError:
            # Handle the case where self.selected_element is not in selector
            pass

    def on_cancel(self):
        selector = []
        QApplication.quit()

    def on_explorer(self):
        self.open_highlight_signal.emit()

    def on_perform(self):
        global final_selector

        selector_field = self.selector_display.toPlainText().strip()
        if not selector_field:
            QMessageBox.warning(self, "Warning", "Selector is not chosen")
            return
        action = self.dropdown.currentText()
        if not action:
            QMessageBox.warning(self, "Warning", "Action is not chosen")
            return
        element = find_element(final_selector)
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
        if not self.highlight_window:
            self.highlight_window = HighlightWindow()
            self.highlight_window.finished.connect(self.on_highlight_finished)
        self.highlight_window.show()

    def on_highlight_finished(self, selector_elements):
        global selector
        if selector_elements:
            selector = selector_elements
        self.showNormal()
        self.properties_table.clearContents()
        self.properties_table.setRowCount(0)
        self.populate_tree()


if __name__ == "__main__":
    final_selector = []
    selector = []
    screen_width, screen_height = pyautogui.size()
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("pythonrpa_logo.ico")))
    window = SelectorExplorer()
    window.show()
    app.exec()
