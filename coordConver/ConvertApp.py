# coding: utf-8
import sys
import os
import pandas as pd
from coordinateConverter import *
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox, 
                             QGroupBox, QMessageBox, QTableWidget, QTableWidgetItem, 
                             QProgressBar, QListWidget, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class ConvertThread(QThread):
    progress_updated = pyqtSignal(int)
    conversion_finished = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, input_files, output_dir, lng_col, lat_col, convert_type, parent=None):
        super().__init__(parent)
        self.input_files = input_files
        self.output_dir = output_dir
        self.lng_col = lng_col
        self.lat_col = lat_col
        self.convert_type = convert_type
        self.running = True

    def run(self):
        try:
            total_files = len(self.input_files)
            for i, input_file in enumerate(self.input_files):
                if not self.running:
                    break
                try:
                    self.progress_updated.emit(int((i + 1) * 100 / total_files))
                    if input_file.endswith('.csv'):
                        df = pd.read_csv(input_file)
                    else:
                        df = pd.read_excel(input_file)

                    converted_coords = []
                    for _, row in df.iterrows():
                        if not self.running:
                            break
                        try:
                            lng_val = row[self.lng_col]
                            lat_val = row[self.lat_col]
                            if pd.isnull(lng_val) or pd.isnull(lat_val):
                                raise ValueError("经纬度为空")
                            lng = float(lng_val)
                            lat = float(lat_val)
                            if self.convert_type == 0:
                                converted = gcj02_to_wgs84(lng, lat)
                            elif self.convert_type == 1:
                                converted = gcj02_to_bd09(lng, lat)
                            elif self.convert_type == 2:
                                converted = bd09_to_gcj02(lng, lat)
                            elif self.convert_type == 3:
                                converted = wgs84_to_gcj02(lng, lat)
                            elif self.convert_type == 4:
                                converted = bd09_to_wgs84(lng, lat)
                            elif self.convert_type == 5:
                                converted = wgs84_to_bd09(lng, lat)
                            converted_coords.append(converted)
                        except Exception as e:
                            converted_coords.append([None, None])
                            print(f"[{os.path.basename(input_file)}] 转换失败: {str(e)}")

                    df[f"{self.lng_col}_converted"] = [coord[0] for coord in converted_coords]
                    df[f"{self.lat_col}_converted"] = [coord[1] for coord in converted_coords]

                    input_filename = os.path.basename(input_file)
                    name, ext = os.path.splitext(input_filename)
                    output_path = os.path.join(self.output_dir, f"converted_{name}{ext}")
                    counter = 1
                    while os.path.exists(output_path):
                        output_path = os.path.join(self.output_dir, f"converted_{counter}_{name}{ext}")
                        counter += 1

                    if input_file.endswith('.csv'):
                        df.to_csv(output_path, index=False)
                    else:
                        output_path = os.path.splitext(output_path)[0] + ".xlsx"
                        df.to_excel(output_path, index=False)

                    self.conversion_finished.emit(f"成功转换: {input_filename}")
                except Exception as e:
                    self.error_occurred.emit(f"处理文件 {os.path.basename(input_file)} 时出错: {str(e)}")
        except Exception as e:
            self.error_occurred.emit(f"转换过程中发生严重错误: {str(e)}")

    def stop(self):
        self.running = False

class CoordinateConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("坐标转换工具 by lmq")
        self.setGeometry(100, 100, 1000, 800)
        self.input_files = []
        self.output_dir = ""
        self.df = None
        self.lng_col = ""
        self.lat_col = ""
        self.convert_thread = None
        self.current_preview_file = None
        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        file_group = QGroupBox("文件设置")
        file_layout = QVBoxLayout()
        batch_layout = QHBoxLayout()
        batch_layout.addWidget(QLabel("批量选择文件:"))
        self.batch_check = QCheckBox("启用批量处理")
        self.batch_check.stateChanged.connect(self.toggle_batch_mode)
        batch_layout.addWidget(self.batch_check)
        file_layout.addLayout(batch_layout)

        self.single_file_layout = QHBoxLayout()
        self.single_file_layout.addWidget(QLabel("输入文件:"))
        self.input_line = QLineEdit()
        self.input_line.setReadOnly(True)
        self.single_file_layout.addWidget(self.input_line)
        input_btn = QPushButton("选择文件")
        input_btn.clicked.connect(self.select_input_file)
        self.single_file_layout.addWidget(input_btn)
        file_layout.addLayout(self.single_file_layout)

        self.batch_file_layout = QHBoxLayout()
        self.batch_file_layout.addWidget(QLabel("批量文件列表:"))
        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self.handle_file_selected)
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        self.batch_file_layout.addWidget(self.file_list)
        batch_btn_layout = QVBoxLayout()
        self.add_files_btn = QPushButton("添加文件")
        self.add_files_btn.clicked.connect(self.add_files)
        batch_btn_layout.addWidget(self.add_files_btn)
        self.remove_files_btn = QPushButton("移除选中")
        self.remove_files_btn.clicked.connect(self.remove_files)
        batch_btn_layout.addWidget(self.remove_files_btn)
        self.clear_files_btn = QPushButton("清空列表")
        self.clear_files_btn.clicked.connect(self.clear_files)
        batch_btn_layout.addWidget(self.clear_files_btn)
        self.batch_file_layout.addLayout(batch_btn_layout)
        file_layout.addLayout(self.batch_file_layout)
        
        # 初始禁用批量相关控件
        self.batch_file_layout.setEnabled(False)
        self.file_list.setEnabled(False)
        self.add_files_btn.setEnabled(False)
        self.remove_files_btn.setEnabled(False)
        self.clear_files_btn.setEnabled(False)

        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_line = QLineEdit()
        self.output_line.setReadOnly(True)
        output_layout.addWidget(self.output_line)
        output_btn = QPushButton("选择目录")
        output_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(output_btn)
        file_layout.addLayout(output_layout)

        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        field_group = QGroupBox("坐标字段选择")
        field_layout = QHBoxLayout()
        field_layout.addWidget(QLabel("经度字段:"))
        self.lng_combo = QComboBox()
        field_layout.addWidget(self.lng_combo)
        field_layout.addWidget(QLabel("纬度字段:"))
        self.lat_combo = QComboBox()
        field_layout.addWidget(self.lat_combo)
        field_group.setLayout(field_layout)
        main_layout.addWidget(field_group)

        convert_group = QGroupBox("转换选项")
        convert_layout = QVBoxLayout()
        self.convert_type = QComboBox()
        self.convert_type.addItems([
            "GCJ02(火星坐标) -> WGS84(GPS坐标)",
            "GCJ02(火星坐标) -> BD09(百度坐标)",
            "BD09(百度坐标) -> GCJ02(火星坐标)",
            "WGS84(GPS坐标) -> GCJ02(火星坐标)",
            "BD09(百度坐标) -> WGS84(GPS坐标)",
            "WGS84(GPS坐标) -> BD09(百度坐标)"
        ])
        convert_layout.addWidget(self.convert_type)
        convert_group.setLayout(convert_layout)
        main_layout.addWidget(convert_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备就绪 (%p%)")
        main_layout.addWidget(self.progress_bar)

        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        main_layout.addWidget(QLabel("数据预览:"))
        main_layout.addWidget(self.preview_table)

        btn_layout = QHBoxLayout()
        preview_btn = QPushButton("预览数据")
        preview_btn.clicked.connect(self.preview_data)
        btn_layout.addWidget(preview_btn)
        self.convert_btn = QPushButton("执行转换")
        self.convert_btn.clicked.connect(self.convert_data)
        btn_layout.addWidget(self.convert_btn)
        self.stop_btn = QPushButton("停止转换")
        self.stop_btn.clicked.connect(self.stop_conversion)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)
        main_layout.addLayout(btn_layout)

        self.log_list = QListWidget()
        main_layout.addWidget(QLabel("转换日志:"))
        main_layout.addWidget(self.log_list)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def toggle_batch_mode(self, state):
        is_batch = state == Qt.Checked
        
        # 切换单文件和批量文件控件的可用状态
        self.single_file_layout.setEnabled(not is_batch)
        self.input_line.setEnabled(not is_batch)
        input_btn = self.single_file_layout.itemAt(2).widget()
        input_btn.setEnabled(not is_batch)
        
        # 设置批量文件列表和相关按钮的可用状态
        self.batch_file_layout.setEnabled(is_batch)
        self.file_list.setEnabled(is_batch)
        self.add_files_btn.setEnabled(is_batch)
        self.remove_files_btn.setEnabled(is_batch)
        self.clear_files_btn.setEnabled(is_batch)
        
        # 清除当前选择和预览
        self.input_line.clear()
        self.file_list.clear()
        self.input_files = []
        self.current_preview_file = None
        self.preview_table.clear()
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        self.reset_progress()
        
        # 重置字段选择
        self.lng_combo.clear()
        self.lat_combo.clear()

    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "支持的文件 (*.xlsx *.xls *.csv)")
        if file_path:
            self.input_line.setText(file_path)
            self.input_files = [file_path]
            self.current_preview_file = file_path  # 先设置当前预览文件
            self.reset_progress()
            try:
                if file_path.endswith('.csv'):
                    self.df = pd.read_csv(file_path, nrows=1)
                else:
                    self.df = pd.read_excel(file_path, nrows=1)
                
                self.lng_combo.clear()
                self.lat_combo.clear()
                
                # 自动识别经纬度字段
                lng_candidates = ['lng', 'longitude', '经度', 'lon', 'x']
                lat_candidates = ['lat', 'latitude', '纬度', 'y']
                
                lng_found = False
                lat_found = False
                
                for col in self.df.columns:
                    self.lng_combo.addItem(col)
                    self.lat_combo.addItem(col)
                    col_lower = str(col).lower()
                    if not lng_found and any(c in col_lower for c in lng_candidates):
                        self.lng_combo.setCurrentText(col)  # 修正拼写错误
                        lng_found = True
                    if not lat_found and any(c in col_lower for c in lat_candidates):
                        self.lat_combo.setCurrentText(col)  # 修正拼写错误
                        lat_found = True
                
                # 立即预览数据
                self.preview_data()
            except Exception as e:
                QMessageBox.warning(self, "错误", f"读取文件失败: {str(e)}")
                self.current_preview_file = None  # 出错时重置预览文件

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_line.setText(dir_path)
            self.output_dir = dir_path

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择多个文件", "", "支持的文件 (*.xlsx *.xls *.csv)")
        if files:
            for file in files:
                if file not in self.input_files:
                    self.input_files.append(file)
                    self.file_list.addItem(os.path.basename(file))
            
            self.reset_progress()
            
            if self.input_files:
                self.current_preview_file = self.input_files[0]
                self.load_file_headers(self.current_preview_file)

    def remove_files(self):
        selected_items = self.file_list.selectedItems()
        rows = sorted([self.file_list.row(item) for item in selected_items], reverse=True)
        for row in rows:
            if 0 <= row < len(self.input_files):
                if self.input_files[row] == self.current_preview_file:
                    self.current_preview_file = None
                del self.input_files[row]
                self.file_list.takeItem(row)
        
        if self.input_files:
            self.current_preview_file = self.input_files[0]
            self.load_file_headers(self.current_preview_file)
        else:
            self.current_preview_file = None
            self.preview_table.clear()
            self.preview_table.setColumnCount(0)
            self.preview_table.setRowCount(0)
        
        self.reset_progress()

    def clear_files(self):
        self.file_list.clear()
        self.input_files = []
        self.current_preview_file = None
        self.preview_table.clear()
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        self.reset_progress()

    def reset_progress(self):
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("准备就绪 (%p%)")

    def load_file_headers(self, file_path):
        try:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path, nrows=1)
            else:
                self.df = pd.read_excel(file_path, nrows=1)
            self.lng_combo.clear()
            self.lat_combo.clear()
            lng_candidates = ['lng', 'longitude', '经度', 'lon', 'x']
            lat_candidates = ['lat', 'latitude', '纬度', 'y']
            lng_found = False
            lat_found = False
            for col in self.df.columns:
                self.lng_combo.addItem(col)
                self.lat_combo.addItem(col)
                col_lower = str(col).lower()
                if not lng_found and any(c in col_lower for c in lng_candidates):
                    self.lng_combo.setCurrentText(col)
                    lng_found = True
                if not lat_found and any(c in col_lower for c in lat_candidates):
                    self.lat_combo.setCurrentText(col)
                    lat_found = True
            self.preview_data()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"读取文件失败: {str(e)}")

    def preview_data(self):
        # 修改检查条件，只需要检查current_preview_file
        if not self.current_preview_file:
            QMessageBox.warning(self, "警告", "请先选择输入文件")
            return
        try:
            if self.current_preview_file.endswith('.csv'):
                self.df = pd.read_csv(self.current_preview_file, nrows=10)
            else:
                self.df = pd.read_excel(self.current_preview_file, nrows=10)
            self.preview_table.setColumnCount(len(self.df.columns))
            self.preview_table.setRowCount(len(self.df))
            self.preview_table.setHorizontalHeaderLabels(self.df.columns)
            for i in range(len(self.df)):
                for j in range(len(self.df.columns)):
                    item = QTableWidgetItem(str(self.df.iloc[i, j]))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                    self.preview_table.setItem(i, j, item)
            self.preview_table.resizeColumnsToContents()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"预览数据失败: {str(e)}")

    def handle_file_selected(self, item):
        filename = item.text()
        for path in self.input_files:
            if os.path.basename(path) == filename:
                self.current_preview_file = path
                self.load_file_headers(path)
                break

    def convert_data(self):
        if not self.input_files:
            QMessageBox.warning(self, "警告", "请先选择输入文件")
            return
        if not self.output_dir:
            QMessageBox.warning(self, "警告", "请先选择输出目录")
            return
        lng_col = self.lng_combo.currentText()
        lat_col = self.lat_combo.currentText()
        if not lng_col or not lat_col:
            QMessageBox.warning(self, "警告", "请选择经度和纬度字段")
            return
        self.convert_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setFormat("转换中... (%p%)")
        self.convert_thread = ConvertThread(
            self.input_files,
            self.output_dir,
            lng_col,
            lat_col,
            self.convert_type.currentIndex()
        )
        self.convert_thread.progress_updated.connect(self.update_progress)
        self.convert_thread.conversion_finished.connect(self.show_conversion_result)
        self.convert_thread.error_occurred.connect(self.show_error)
        self.convert_thread.finished.connect(self.conversion_complete)
        self.convert_thread.start()

    def stop_conversion(self):
        if self.convert_thread and self.convert_thread.isRunning():
            self.convert_thread.stop()
            self.convert_thread.wait()
            QMessageBox.information(self, "信息", "转换已停止")
            self.reset_progress()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def show_conversion_result(self, message):
        self.log_list.addItem(f"✅ {message}")

    def show_error(self, error_message):
        self.log_list.addItem(f"❌ {error_message}")

    def conversion_complete(self):
        self.convert_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.convert_thread = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CoordinateConverterApp()
    window.show()
    sys.exit(app.exec_())