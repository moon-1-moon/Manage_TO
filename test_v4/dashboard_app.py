import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QTableWidget, QTableWidgetItem, QSplitter, 
    QMessageBox, QInputDialog, QFileDialog, QHeaderView
)
from PyQt6.QtCore import Qt
from data_manager import DataManager
from ui_components import SlicerPanel, VizPanel

class DashboardApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("정원 관리 프로그램")
        self.resize(1200, 800)
        
        self.data_manager = DataManager()
        self.data_manager.load_data()
        
        self.current_filters = {} # {col: [values]}
        
        self.init_ui()
        self.refresh_all()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        
        # --- Sidebar (Slicers + Mode Switch) ---
        sidebar_layout = QVBoxLayout()
        
        # Mode Switch Buttons
        mode_layout = QHBoxLayout()
        self.btn_std = QPushButton("기준정원")
        self.btn_ops = QPushButton("운영정원")
        self.btn_std.setCheckable(True)
        self.btn_ops.setCheckable(True)
        self.btn_std.clicked.connect(lambda: self.switch_mode("STD"))
        self.btn_ops.clicked.connect(lambda: self.switch_mode("OPS"))
        
        mode_layout.addWidget(self.btn_std)
        mode_layout.addWidget(self.btn_ops)
        sidebar_layout.addLayout(mode_layout)
        
        # Slicers
        self.slicer_panel = SlicerPanel(self.data_manager.filter_cols)
        self.slicer_panel.filter_changed.connect(self.on_filter_changed)
        sidebar_layout.addWidget(self.slicer_panel)
        
        main_layout.addLayout(sidebar_layout)
        
        # --- Main Content (Viz + Grid) ---
        content_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Top: Visualization
        self.viz_panel = VizPanel()
        content_splitter.addWidget(self.viz_panel)
        content_splitter.setStretchFactor(0, 1) # 40% height ish
        
        # Bottom: Grid + CRUD
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        
        # Grid
        self.table = QTableWidget()
        bottom_layout.addWidget(self.table)
        
        # CRUD Buttons
        btn_layout = QHBoxLayout()
        btn_add = QPushButton("행 추가")
        btn_add.clicked.connect(self.add_row)
        btn_del = QPushButton("행 삭제")
        btn_del.clicked.connect(self.delete_row)
        btn_save = QPushButton("엑셀 내보내기")
        btn_save.clicked.connect(self.export_excel)
        
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_del)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        bottom_layout.addLayout(btn_layout)
        
        content_splitter.addWidget(bottom_widget)
        content_splitter.setStretchFactor(1, 2)
        
        main_layout.addWidget(content_splitter)
        
        # Connect grid changes
        self.table.itemChanged.connect(self.on_item_changed)
        
        # Initial State
        self.switch_mode("STD")

    def switch_mode(self, mode):
        self.data_manager.set_current_type(mode)
        
        if mode == "STD":
            self.btn_std.setChecked(True)
            self.btn_ops.setChecked(False)
            self.btn_std.setStyleSheet("background-color: #4CAF50; color: white;")
            self.btn_ops.setStyleSheet("")
        else:
            self.btn_std.setChecked(False)
            self.btn_ops.setChecked(True)
            self.btn_ops.setStyleSheet("background-color: #2196F3; color: white;")
            self.btn_std.setStyleSheet("")
            
        # Reset filters when switching? 
        # User might want to keep filters if columns are same. 
        # But for safety let's keep them if compatible.
        self.refresh_all()

    def on_filter_changed(self, col, selected_values):
        self.current_filters[col] = selected_values
        self.refresh_all(update_slicers=False)
        # Update OTHER slicers dynamically
        self.update_dynamic_slicers(expert_col=col)

    def update_dynamic_slicers(self, expert_col=None):
        # Determine valid options for other slicers based on current filtering
        for col in self.data_manager.filter_cols:
            if col == expert_col: continue # Don't filter the one user just clicked
            
            # Temporary filter excluding this column to see what's available
            # Actually, standard behavior:
            # 1. Available options in Slicer B are constrained by Selection in Slicer A
            # 2. If I select "Busan" in Dept1, Dept2 should only show Depts in Busan.
            
            options = self.data_manager.get_unique_values(col, self.current_filters)
            self.slicer_panel.update_slicer_options(col, options)

    def refresh_all(self, update_slicers=True):
        # 1. Get Data
        df = self.data_manager.get_filtered_data(self.current_filters)
        
        # 2. Update Viz
        self.viz_panel.update_charts(df, self.current_filters)
        
        # 3. Update Grid
        self.update_table(df)
        
        # 4. Update Slicers (Initial Load or Mode Switch)
        if update_slicers:
            self.update_dynamic_slicers()

    def update_table(self, df):
        self.table.blockSignals(True)
        self.table.clear()
        
        if df.empty:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.table.blockSignals(False)
            return
            
        cols = df.columns.tolist()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        
        for r, row in enumerate(df.itertuples(index=False)):
            for c, val in enumerate(row):
                str_val = str(val)
                if str_val == "nan":
                    str_val = ""
                item = QTableWidgetItem(str_val)
                self.table.setItem(r, c, item)
        
        self.table.blockSignals(False)

    def on_item_changed(self, item):
        row = item.row()
        col = item.column()
        # Note: This is tricky with filtering. 
        # If we filter, the visual row 0 is not necessarily original row 0.
        # DataManager needs to handle edits based on *filtered* view or we need ID.
        # For this MVP, let's assume raw index mapping if possible, 
        # BUT DataFrame index is preserved in get_filtered_data!
        # So we should use the DataFrame's index.
        
        # However, QTableWidget doesn't store the original index automatically.
        # We need to rely on the fact that `update_table` iterates over `df`.
        # So visual row `r` corresponds to `df.iloc[r]`. 
        # `df` is the filtered one.
        
        current_df = self.data_manager.get_filtered_data(self.current_filters)
        original_idx = current_df.index[row]
        col_name = current_df.columns[col]
        
        val = item.text()
        self.data_manager.update_row(original_idx, col_name, val)
        
        # Trigger explicit update if needed (e.g. if Department changed, it might disappear from filter)
        # But avoid infinite loops.

    def add_row(self):
        # Simple dialog to insert at top or bottom? Or just add empty row
        # User requirement: "데이터 행 삽입 기능(삽입 위치지정 포함)"
        rows = self.table.rowCount()
        idx, ok = QInputDialog.getInt(self, "행 삽입", "삽입할 위치(행 번호):", 0, 0, rows+1)
        if ok:
            # Create a dummy row matching columns
            cols = self.data_manager.get_current_df().columns
            new_data = {c: "" for c in cols}
            new_data['정원'] = 0
            
            # We need to map visual index to real index if possible, 
            # but usually insertion on filtered data is ambiguous. 
            # We will insert into the *main* dataframe at the approximate position?
            # Or just append. 
            # For simplicity: Append to Main DF, or insert at index of Main DF.
            # Let's just append for MVP or insert at 0.
            
            self.data_manager.add_row(new_data, idx) 
            self.refresh_all()

    def delete_row(self):
        selected_rows = sorted(set(idx.row() for idx in self.table.selectedIndexes()), reverse=True)
        if not selected_rows:
            QMessageBox.warning(self, "경고", "삭제할 행을 선택하세요.")
            return

        confirm = QMessageBox.question(self, "삭제 확인", f"{len(selected_rows)}개 행을 삭제하시겠습니까?", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if confirm == QMessageBox.StandardButton.Yes:
            # Map visual rows to real indices
            current_df = self.data_manager.get_filtered_data(self.current_filters)
            real_indices = [current_df.index[r] for r in selected_rows]
            
            self.data_manager.delete_rows(real_indices)
            self.refresh_all()

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "엑셀 내보내기", "", "Excel Files (*.xlsx)")
        if path:
            success, msg = self.data_manager.export_data(path)
            if success:
                QMessageBox.information(self, "성공", "파일이 저장되었습니다.")
            else:
                QMessageBox.critical(self, "오류", f"저장 실패: {msg}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Optional: Set global font
    font = app.font()
    font.setFamily("Malgun Gothic")
    font.setPointSize(10)
    app.setFont(font)
    
    window = DashboardApp()
    window.show()
    sys.exit(app.exec())
