from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QCheckBox, QHeaderView, QLabel, QTextEdit, QAbstractItemView, QPushButton, QMessageBox
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QFontMetrics, QPainter, QPen, QIcon
import sys
import time
from pywinauto import uia_element_info
from pywinauto.uia_element_info import UIAElementInfo
from pywinauto.controls.uiawrapper import UIAWrapper
import json
import os
import win32process
import win32gui
import pyautogui
import copy
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
    return None


def find_matches(parent, selector, level=0):
    if level >= len(selector):
        return [parent]
    matches = []
    expected_props = selector[level]
    if not hasattr(parent, "children"):
        print(f"Skipping invalid parent at level {level}: {parent}")
        return []
    try:
        children = parent.children()
    except AttributeError:
        print(f"Error: Parent at level {level} does not support 'children()': {parent}")
        return []
    if expected_props.get("ctrl_index"):
        try:
            matched_descendants = find_matches(children[expected_props.get("ctrl_index")], selector, level + 1)
            matches.extend(matched_descendants)
            return matches
        except Exception as e:
            return []

    def selector_params(param: str, child):
        if param == "class_name":
            value = child.class_name
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "title":
            value = child.name
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "control_type":
            value = child.control_type
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "rich_text":
            value = child.rich_text
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "visible":
            value = child.visible
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "enabled":
            value = child.enabled
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "control_id":
            value = child.control_id
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "automation_id":
            value = child.automation_id
            if value == "" or value == 0 or value == {} or value == []:
                value = None
            return value
        elif param == "rectangle":
            value = {"left": child.rectangle.left, "right": child.rectangle.right, "top": child.rectangle.top, "bottom": child.rectangle.bottom}
            if value == "" or value == "0" or value == {} or value == []:
                value = None
            return value

    for child in children:
        element_found = True
        for key, _ in expected_props.items():
            selector_key_value = expected_props.get(key)
            if selector_key_value == "" or selector_key_value == 0 or selector_key_value == {} or selector_key_value == []:
                selector_key_value = None
            if selector_key_value != selector_params(key, child):
                element_found = False
        if element_found:
            matched_descendants = find_matches(child, selector, level + 1)
            matches.extend(matched_descendants)
    return matches


def find_elements(selector):
    if not selector or not isinstance(selector[0], dict):
        print("Invalid selector:", selector)
        return []
    desktop = UIAElementInfo()
    search_args = {k: v for k, v in selector[0].items() if v and (k == "title" or k == "class_name")}
    window = desktop.children(**search_args)
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
        print(f"COMError in CompareElements: {e}")
        return False  # Handle invalid elements safely


def wait_elements_to_appear(selectors_list: list[list], wait_time: int=30):
    start_time = time.time()
    found_list = []
    for selector in selectors_list:
        while True:
            result = find_elements(selector)
            if result:
                found_list.append(True)
                break
            if time.time() - start_time > wait_time:
                break
    if len(found_list) == len(selectors_list) and False not in found_list:
        return True
    else:
        return False


def wait_elements_to_disappear(selectors_list: list[list], wait_time: int=30):
    start_time = time.time()
    found_list = []
    for selector in selectors_list:
        while True:
            result = find_elements(selector)
            if not result:
                found_list.append(True)
                break
            if time.time() - start_time > wait_time:
                break
    if len(found_list) == len(selectors_list) and False not in found_list:
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

        # **Find the actual index of the element among siblings**
        try:
            return siblings.index(element)
        except ValueError:
            return None  # If element is not found for some reason


    def get_element_properties(self, element):
        """Extracts all available properties of a UIA element into a dictionary."""
        properties = {
            "class_name": element.class_name,
            "title": element.name,
            "control_type": element.control_type,
            "rich_text": element.rich_text,
            "visible": element.visible,
            "enabled": element.enabled,
            "control_id": element.control_id,
            "automation_id": element.automation_id,
            "rectangle": {
                "left": element.rectangle.left,
                "right": element.rectangle.right,
                "top": element.rectangle.top,
                "bottom": element.rectangle.bottom
                }
        }
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
            print(e)
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

        # Layouts
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        main_layout.addLayout(top_layout)

        # Left Panel (Tree View) with Scrollbars
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Elements Tree")
        self.tree.itemClicked.connect(self.on_element_selected)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.tree.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)  # Smooth scrolling
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)  # Allow user resizing
        self.tree.setIndentation(20)  # Proper indentation for child elements
        top_layout.addWidget(self.tree, 3)

        # Right Panel (Properties Table) with Scrollbars
        self.properties_table = QTableWidget()
        self.properties_table.setColumnCount(2)
        self.properties_table.setHorizontalHeaderLabels(["", "Property"])
        self.properties_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.properties_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.properties_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.properties_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        top_layout.addWidget(self.properties_table, 3)

        # Bottom Panel (Selector Display) with Scrollbars
        self.selector_display = QTextEdit()
        self.selector_display.setReadOnly(True)
        self.selector_display.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.selector_display.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.selector_display.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        main_layout.addWidget(QLabel("Generated Selector:"))
        main_layout.addWidget(self.selector_display)

        button_layout = QHBoxLayout()
        main_layout.addLayout(button_layout)

        # Submit Button
        self.submit_button = QPushButton("Submit")
        self.submit_button.clicked.connect(self.on_submit)
        button_layout.addWidget(self.submit_button)

        # Cancel Button
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.on_cancel)
        button_layout.addWidget(self.cancel_button)

        # Explore Button
        self.explore_button = QPushButton("Explore")
        self.explore_button.clicked.connect(self.on_explorer)
        button_layout.addWidget(self.explore_button)

        # Click on element Button
        self.click_on_button = QPushButton("Click")
        self.click_on_button.clicked.connect(self.on_click)
        button_layout.addWidget(self.click_on_button)

        self.properties_table.clearContents()
        self.properties_table.setRowCount(0)

        # Populate tree
        self.populate_tree()

    def populate_tree(self):
        self.tree.clear()
        global selector
        item_map = {}  # Stores elements by their index for parent-child linking

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
        self.adjust_column_width()

    def get_item_depth(self, item):
        """Calculate depth (hierarchy level) of an item by counting its parents."""
        depth = 0
        while item.parent():
            item = item.parent()
            depth += 1
        return depth


    def adjust_column_width(self):
        font_metrics = QFontMetrics(self.tree.font())

        # Ensure there are items before computing width
        tree_items = self.iterate_tree_items()

        if not tree_items:  # Avoid max() on empty sequence
            return

        # Compute max width based on longest text + indentation level
        max_width = max(
            font_metrics.horizontalAdvance(item.text(0)) + (self.tree.indentation() * self.get_item_depth(item))
            for item in tree_items
        )

        max_width  = max_width + int(max_width/10)

        # 🔥 Set the new column width
        self.tree.setColumnWidth(0, max_width)


    def iterate_tree_items(self):
        def get_all_items(parent):
            for i in range(parent.childCount()):
                child = parent.child(i)
                yield child
                yield from get_all_items(child)

        return [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())] + list(
            get_all_items(self.tree.invisibleRootItem()))

    def on_element_selected(self, item, column):
        self.selected_element = item.data(0, 1)
        if self.selected_element:
            self.update_properties_table()

    def create_checkbox_callback(self, key):
        """Returns a proper callback function to prevent lambda closure issues."""
        return lambda state, key_copy=key: self.on_checkbox_state_changed(key_copy, state)

    def update_properties_table(self):
        """Updates the properties table when an element is selected."""
        if not self.selected_element:
            return

        self.properties_table.clearContents()
        row_counts = len(self.selected_element) - 1
        if self.selected_element.get("level") == 0:
            row_counts = 2
        self.properties_table.setRowCount(row_counts)
        row = 0
        for key, value in self.selected_element.items():
            if key == "level":
                continue

            if self.selected_element.get("level") == 0:
                if key != "title" and key != "class_name":
                    continue

            checkbox = QCheckBox()
            level = self.selected_element.get("level")
            if level is not None:
                # Check if the key is in the checkboxes dictionary for this level
                checked = key in self.checkboxes.get(level, [])
                checkbox.setChecked(checked)
                # Connect the checkbox state change to the callback
                checkbox.stateChanged.connect(self.create_checkbox_callback(key))

            self.properties_table.setCellWidget(row, 0, checkbox)
            self.properties_table.setItem(row, 1, QTableWidgetItem(f"{key}: {value}"))
            row += 1

        self.update_selector()

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
            selector_ = copy.deepcopy(selector_text)
            final_selector = selector_

            # Update the bottom panel
            self.selector_display.clear()
            self.selector_display.setText(json.dumps(selector_, ensure_ascii=False, indent=4))
        except ValueError:
            # Handle the case where self.selected_element is not in selector
            print("Selected element not found in selector!")


    def on_submit(self):
        global final_selector
        text = self.selector_display.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "Warning", "Selector is not chosen")
        else:
            QApplication.quit()

    def on_cancel(self):
        selector = []
        QApplication.quit()

    def on_explorer(self):
        self.open_highlight_signal.emit()

    def on_click(self):
        global final_selector
        text = self.selector_display.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "Warning", "Selector is not chosen")
        else:
            result = find_element(final_selector)
            if result:
                self.hide()
                find_element(final_selector).click_input()
                self.show()
            else:
                QMessageBox.warning(self, "Warning", "Selector is not found")

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
    global final_selector
    global selector
    selector = []  # Ensure selector is a global list
    screen_width, screen_height = pyautogui.size()
    app = QApplication(sys.argv)  # Create QApplication once
    app.setWindowIcon(QIcon(resource_path("pythonrpa_logo.ico")))
    window = SelectorExplorer()  # Start with `SelectorExplorer`
    window.show()
    app.exec()  # Start event loop
