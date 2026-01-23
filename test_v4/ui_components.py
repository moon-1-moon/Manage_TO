from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QAbstractItemView,
    QGroupBox, QScrollArea, QFrame, QPushButton, QMenu, QWidgetAction, QListWidgetItem
)
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from PyQt6.QtCore import pyqtSignal, Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Set font for matplotlib to support Korean
plt.rcParams['font.family'] = 'Malgun Gothic'

class MultiSelectButton(QPushButton):
    selection_changed = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setText("전체")
        self.setStyleSheet("""
            QPushButton {
                text-align: left;
                padding: 5px;
                border: 1px solid #aaa;
                border-radius: 3px;
                background-color: white;
            }
            QPushButton::menu-indicator {
                image: none; /* Hide default indicator to use custom logic if needed, or keep it */
                subcontrol-position: right center;
                subcontrol-origin: padding;
                width: 20px;
            }
        """)
        
        # Menu with ListWidget
        self.menu = QMenu(self)
        self.menu.setStyleSheet("QMenu { background-color: white; border: 1px solid #aaa; }")
        
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection) # Handle checks manually
        self.list_widget.itemChanged.connect(self.on_item_changed)
        
        # Container to hold list widget (needed for QWidgetAction)
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.list_widget)
        
        # Minimum size for the list
        self.list_widget.setMinimumHeight(200)
        self.list_widget.setMinimumWidth(200)
        
        self.action = QWidgetAction(self.menu)
        self.action.setDefaultWidget(container)
        self.menu.addAction(self.action)
        
        self.setMenu(self.menu)

    def showEvent(self, event):
        super().showEvent(event)
        # Ensure list widget width matches button or is reasonable
        self.list_widget.setMinimumWidth(self.width())

    def on_item_changed(self, item):
        self.update_text()
        self.selection_changed.emit(self.get_checked_items())

    def update_text(self):
        items = self.get_checked_items()
        if not items:
            self.setText("전체")
        elif len(items) == self.list_widget.count():
            self.setText("전체")
        else:
            self.setText(", ".join(items))

    def get_checked_items(self):
        checked = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                checked.append(item.text())
        return checked

    def add_items(self, items, selected_items=None):
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        
        if selected_items is None:
            selected_items = []
            
        for text in items:
            item = QListWidgetItem(text)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            
            # Logic: If 'selected_items' is empty, it means "All" (conceptually in filter),
            # BUT visually in the checkbox list, should they be checked or unchecked?
            # Standard pattern: Unchecked means "All" in the current implementation (get_filtered_data checks 'if selected_vals:').
            # So if we want "All", we ensure nothing is checked.
            
            if text in selected_items:
                item.setCheckState(Qt.CheckState.Checked)
            else:
                item.setCheckState(Qt.CheckState.Unchecked)
                
            self.list_widget.addItem(item)
            
        self.update_text()
        self.list_widget.blockSignals(False)
    
    def clear_items(self):
        self.list_widget.clear()
        self.setText("전체")


class SlicerPanel(QScrollArea):
    filter_changed = pyqtSignal(str, list) # col_name, selected_values

    def __init__(self, filter_cols):
        super().__init__()
        self.setWidgetResizable(True)
        self.filter_cols = filter_cols
        self.slicer_widgets = {}
        
        container = QWidget()
        self.layout = QVBoxLayout(container)
        self.layout.setSpacing(15)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        for col in self.filter_cols:
            # Container for Label + ComboBox
            vbox = QVBoxLayout()
            vbox.setSpacing(5)
            
            label = QLabel(col)
            label.setStyleSheet("font-weight: bold; color: #333;")
            
            combo = MultiSelectButton()
            combo.selection_changed.connect(lambda items, c=col: self.on_selection_change(c, items))
            
            vbox.addWidget(label)
            vbox.addWidget(combo)
            
            self.layout.addLayout(vbox)
            self.slicer_widgets[col] = combo
            
        self.setWidget(container)
        self.setFixedWidth(220)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setStyleSheet("QScrollArea { background-color: #f5f5f5; border: none; } QWidget { background-color: #f5f5f5; }")

    def on_selection_change(self, col, items):
        self.filter_changed.emit(col, items)

    def update_slicer_options(self, col, options):
        if col in self.slicer_widgets:
            widget = self.slicer_widgets[col]
            current_selection = widget.get_checked_items()
            widget.add_items(options, current_selection)

class VizPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout(self)
        
        # Total Count
        self.total_label = QLabel("총 정원: 0명")
        self.total_label.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px; background-color: white; border-bottom: 2px solid #ddd;")
        self.layout.addWidget(self.total_label)
        
        # Charts Area
        charts_layout = QHBoxLayout()
        
        # Dept Chart
        self.dept_fig = Figure(figsize=(5, 4), dpi=100)
        self.dept_canvas = FigureCanvas(self.dept_fig)
        self.dept_ax = self.dept_fig.add_subplot(111)
        charts_layout.addWidget(self.dept_canvas)
        
        # Rank Chart (Job Grade)
        self.rank_fig = Figure(figsize=(5, 4), dpi=100)
        self.rank_canvas = FigureCanvas(self.rank_fig)
        self.rank_ax = self.rank_fig.add_subplot(111)
        charts_layout.addWidget(self.rank_canvas)
        
        self.layout.addLayout(charts_layout)

    def update_charts(self, df, current_filters=None):
        if df.empty:
            self.total_label.setText("총 정원: 0명")
            self.dept_ax.clear()
            self.rank_ax.clear()
            self.dept_canvas.draw()
            self.rank_canvas.draw()
            return
            
        total_count = df['정원'].sum()
        self.total_label.setText(f"총 정원: {total_count:,.1f}명")
        
        # --- Dept Chart Logic ---
        # Determine drill-down level
        dept_col = "부서1"
        if current_filters:
            # Check levels 1 to 5 to decide what to show
            for i in range(1, 6):
                curr_level = f"부서{i}"
                next_level = f"부서{i+1}"
                
                # If current level is filtered (has selection), we try to show next level
                if curr_level in current_filters and current_filters[curr_level]:
                    # Also verify next level exists in DF columns
                    if next_level in df.columns:
                        dept_col = next_level
                    else:
                        dept_col = curr_level # Fallback
                        break
        
        self.dept_ax.clear()
        if dept_col in df.columns:
            # Filter out empty values/nans for better chart
            dept_data = df.groupby(dept_col)['정원'].sum().sort_values(ascending=True)
            # Limit to top N if too many? For now show all.
            dept_data.plot(kind='barh', ax=self.dept_ax, color='skyblue')
            self.dept_ax.set_title(f'{dept_col}별 정원')
            self.dept_ax.set_xlabel('정원')
            # Add labels
            if self.dept_ax.containers:
                self.dept_ax.bar_label(self.dept_ax.containers[0], fmt='%.0f', padding=3)
        
        # --- Rank Chart ---
        self.rank_ax.clear()
        target_col = '직급' if '직급' in df.columns else '직렬' 
        if target_col in df.columns:
            rank_data = df.groupby(target_col)['정원'].sum().sort_values(ascending=True)
            rank_data.plot(kind='barh', ax=self.rank_ax, color='lightgreen')
            self.rank_ax.set_title(f'{target_col}별 정원')
            self.rank_ax.set_xlabel('정원')
            # Add labels
            if self.rank_ax.containers:
                self.rank_ax.bar_label(self.rank_ax.containers[0], fmt='%.0f', padding=3)

        self.dept_fig.tight_layout()
        self.rank_fig.tight_layout()
        self.dept_canvas.draw()
        self.rank_canvas.draw()
