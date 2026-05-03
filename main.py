import sys
import os
import re
from datetime import datetime
from typing import List, Dict, Any

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QLineEdit, QTableWidget,
    QTableWidgetItem, QGroupBox, QComboBox, QSpinBox, QDoubleSpinBox,
    QCheckBox, QFileDialog, QMessageBox, QProgressBar, QSplitter,
    QScrollArea, QFrame, QSizePolicy, QHeaderView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QImage

from config import APP_NAME, APP_VERSION, OUTPUT_DIR, TEMPLATE_DIR
from document_generator import DocumentGenerator, BatchDocumentFormatter
from excel_reader import ExcelReader
from database import DatabaseManager
from chart_generator import ChartGenerator


class WorkerThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    
    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
    
    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class GenerateTab(QWidget):
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.template_path = ""
        self.excel_path = ""
        self.excel_data = []
        self.placeholders = []
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        template_group = QGroupBox("1. 选择模板文件")
        template_layout = QHBoxLayout(template_group)
        
        self.template_edit = QLineEdit()
        self.template_edit.setReadOnly(True)
        self.template_edit.setPlaceholderText("请选择Word模板文件 (.docx)")
        
        self.template_btn = QPushButton("浏览...")
        self.template_btn.clicked.connect(self.select_template)
        
        template_layout.addWidget(self.template_edit)
        template_layout.addWidget(self.template_btn)
        
        layout.addWidget(template_group)
        
        data_group = QGroupBox("2. 选择数据来源")
        data_layout = QVBoxLayout(data_group)
        
        data_source_layout = QHBoxLayout()
        self.data_source_combo = QComboBox()
        self.data_source_combo.addItems(["Excel表格导入", "手动输入单条数据"])
        self.data_source_combo.currentIndexChanged.connect(self.switch_data_source)
        
        data_source_layout.addWidget(QLabel("数据来源:"))
        data_source_layout.addWidget(self.data_source_combo)
        data_source_layout.addStretch()
        
        data_layout.addLayout(data_source_layout)
        
        self.excel_widget = QWidget()
        excel_layout = QHBoxLayout(self.excel_widget)
        excel_layout.setContentsMargins(0, 0, 0, 0)
        
        self.excel_edit = QLineEdit()
        self.excel_edit.setReadOnly(True)
        self.excel_edit.setPlaceholderText("请选择Excel文件 (.xlsx)")
        
        self.excel_btn = QPushButton("浏览...")
        self.excel_btn.clicked.connect(self.select_excel)
        
        self.load_excel_btn = QPushButton("加载数据")
        self.load_excel_btn.clicked.connect(self.load_excel_data)
        self.load_excel_btn.setEnabled(False)
        
        excel_layout.addWidget(self.excel_edit)
        excel_layout.addWidget(self.excel_btn)
        excel_layout.addWidget(self.load_excel_btn)
        
        data_layout.addWidget(self.excel_widget)
        
        self.manual_widget = QWidget()
        manual_layout = QVBoxLayout(self.manual_widget)
        manual_layout.setContentsMargins(0, 0, 0, 0)
        
        self.placeholder_table = QTableWidget()
        self.placeholder_table.setColumnCount(2)
        self.placeholder_table.setHorizontalHeaderLabels(["占位符", "值"])
        self.placeholder_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        manual_layout.addWidget(self.placeholder_table)
        
        self.manual_widget.hide()
        data_layout.addWidget(self.manual_widget)
        
        layout.addWidget(data_group)
        
        output_group = QGroupBox("3. 输出设置")
        output_layout = QVBoxLayout(output_group)
        
        filename_layout = QHBoxLayout()
        self.filename_combo = QComboBox()
        self.filename_combo.addItems(["使用序号命名 (文档_1, 文档_2...)"])
        self.filename_combo.setEnabled(False)
        
        filename_layout.addWidget(QLabel("文件名规则:"))
        filename_layout.addWidget(self.filename_combo)
        filename_layout.addStretch()
        
        output_layout.addLayout(filename_layout)
        
        output_path_layout = QHBoxLayout()
        self.output_edit = QLineEdit(OUTPUT_DIR)
        self.output_edit.setPlaceholderText("输出目录")
        
        self.output_btn = QPushButton("浏览...")
        self.output_btn.clicked.connect(self.select_output_dir)
        
        output_path_layout.addWidget(QLabel("输出目录:"))
        output_path_layout.addWidget(self.output_edit)
        output_path_layout.addWidget(self.output_btn)
        
        output_layout.addLayout(output_path_layout)
        
        layout.addWidget(output_group)
        
        progress_group = QGroupBox("操作进度")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #666;")
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        
        layout.addWidget(progress_group)
        
        btn_layout = QHBoxLayout()
        
        self.preview_btn = QPushButton("预览模板占位符")
        self.preview_btn.clicked.connect(self.preview_placeholders)
        self.preview_btn.setEnabled(False)
        
        self.generate_btn = QPushButton("生成文档")
        self.generate_btn.clicked.connect(self.generate_documents)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        
        btn_layout.addWidget(self.preview_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.generate_btn)
        
        layout.addLayout(btn_layout)
    
    def select_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", TEMPLATE_DIR,
            "Word文档 (*.docx);;所有文件 (*)"
        )
        if file_path:
            self.template_path = file_path
            self.template_edit.setText(file_path)
            self.preview_btn.setEnabled(True)
            self._load_template_placeholders()
    
    def _load_template_placeholders(self):
        try:
            generator = DocumentGenerator(self.template_path)
            self.placeholders = generator.get_placeholders()
            
            self.filename_combo.clear()
            self.filename_combo.addItem("使用序号命名 (文档_1, 文档_2...)")
            for placeholder in self.placeholders:
                self.filename_combo.addItem(f"使用[{placeholder}]作为文件名")
            
            self.generate_btn.setEnabled(True)
            
            if not self.placeholders:
                QMessageBox.information(
                    self, "模板加载完成",
                    f"✅ 模板已成功加载\n\n"
                    f"⚠️ 注意：当前模板中没有检测到占位符\n\n"
                    f"📝 如果需要自动替换内容，请在Word模板中使用以下格式：\n"
                    f"   - 姓名：{{姓名}}\n"
                    f"   - 电话：{{手机号}}\n"
                    f"   - 日期：{{日期}}\n\n"
                    f"💡 使用 {{占位符名}} 格式标记需要替换的位置，然后重新选择模板。"
                )
            else:
                QMessageBox.information(
                    self, "模板加载成功",
                    f"✅ 成功加载模板\n\n"
                    f"📋 检测到 {len(self.placeholders)} 个占位符：\n\n"
                    + "\n".join([f"  • {{{ph}}}" for ph in self.placeholders]) +
                    f"\n\n💡 现在可以选择数据来源（Excel导入或手动输入）进行批量生成！"
                )
        
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法读取模板: {str(e)}")
    
    def preview_placeholders(self):
        if not self.placeholders:
            QMessageBox.information(self, "提示", "模板中没有检测到占位符")
            return
        
        msg = f"检测到 {len(self.placeholders)} 个占位符:\n\n"
        for i, ph in enumerate(self.placeholders, 1):
            msg += f"{i}. {{ {ph} }}\n"
        
        QMessageBox.information(self, "模板占位符", msg)
    
    def switch_data_source(self, index):
        if index == 0:
            self.excel_widget.show()
            self.manual_widget.hide()
        else:
            if not self.template_path:
                QMessageBox.information(
                    self, "使用提示",
                    "📋 使用手动输入功能前，请先完成以下步骤：\n\n"
                    "1. 点击上方的【浏览...】按钮选择一个Word模板文件\n"
                    "2. 模板文件中需要使用 {占位符名} 格式标记需要替换的内容\n\n"
                    "📝 模板格式示例：\n"
                    "   - 姓名：{姓名}\n"
                    "   - 联系电话：{手机号}\n"
                    "   - 签订日期：{日期}\n\n"
                    "💡 提示：先选择模板，程序会自动识别模板中的占位符！"
                )
                self.data_source_combo.setCurrentIndex(0)
                return
            
            if not self.placeholders:
                QMessageBox.warning(
                    self, "提示",
                    "当前选中的模板中没有检测到占位符。\n\n"
                    "请在Word模板中使用 {占位符名} 格式标记需要替换的内容。\n\n"
                    "例如：\n"
                    "  - 姓名：{姓名}\n"
                    "  - 联系电话：{手机号}"
                )
                self.data_source_combo.setCurrentIndex(0)
                return
            
            self.excel_widget.hide()
            self.manual_widget.show()
            self._update_manual_table()
    
    def _update_manual_table(self):
        if not self.placeholders:
            self.placeholder_table.setRowCount(0)
            return
        
        self.placeholder_table.setRowCount(len(self.placeholders))
        for i, ph in enumerate(self.placeholders):
            placeholder_item = QTableWidgetItem(ph)
            placeholder_item.setFlags(placeholder_item.flags() & ~Qt.ItemIsEditable)
            self.placeholder_table.setItem(i, 0, placeholder_item)
            self.placeholder_table.setItem(i, 1, QTableWidgetItem(""))
    
    def select_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if file_path:
            self.excel_path = file_path
            self.excel_edit.setText(file_path)
            self.load_excel_btn.setEnabled(True)
    
    def load_excel_data(self):
        if not self.excel_path:
            return
        
        try:
            reader = ExcelReader(self.excel_path)
            self.excel_data = reader.read_all_data()
            headers = reader.get_headers()
            reader.close()
            
            if self.placeholders:
                self.filename_combo.clear()
                self.filename_combo.addItem("使用序号命名 (文档_1, 文档_2...)")
                for header in headers:
                    self.filename_combo.addItem(f"使用[{header}]作为文件名")
            
            self.filename_combo.setEnabled(True)
            
            QMessageBox.information(
                self, "成功",
                f"成功加载 {len(self.excel_data)} 条数据\n\n"
                f"列名: {', '.join(headers)}"
            )
        
        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载Excel失败: {str(e)}")
    
    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_edit.text())
        if dir_path:
            self.output_edit.setText(dir_path)
    
    def get_key_field(self) -> str:
        current_text = self.filename_combo.currentText()
        match = re.search(r'\[(.*?)\]', current_text)
        if match:
            return match.group(1)
        return None
    
    def get_manual_data(self) -> Dict[str, str]:
        data = {}
        for i in range(self.placeholder_table.rowCount()):
            placeholder_item = self.placeholder_table.item(i, 0)
            if not placeholder_item:
                continue
            
            placeholder = placeholder_item.text()
            value_item = self.placeholder_table.item(i, 1)
            value = value_item.text().strip() if value_item else ""
            data[placeholder] = value
        return data
    
    def generate_documents(self):
        if not self.template_path:
            QMessageBox.warning(self, "提示", "请先选择模板文件")
            return
        
        output_dir = self.output_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "提示", "请选择输出目录")
            return
        
        data_source = self.data_source_combo.currentIndex()
        
        if data_source == 0:
            if not self.excel_data:
                QMessageBox.warning(self, "提示", "请先加载Excel数据")
                return
            data_list = self.excel_data
        else:
            manual_data = self.get_manual_data()
            if not any(manual_data.values()):
                QMessageBox.warning(self, "提示", "请输入占位符的值")
                return
            data_list = [manual_data]
        
        key_field = self.get_key_field()
        
        self.progress_bar.setMaximum(len(data_list))
        self.progress_bar.setValue(0)
        self.status_label.setText("正在生成文档...")
        self.generate_btn.setEnabled(False)
        
        try:
            generator = DocumentGenerator(self.template_path)
            
            success_count = 0
            error_count = 0
            results = []
            
            for i, data in enumerate(data_list):
                try:
                    if key_field and key_field in data and data[key_field]:
                        safe_filename = re.sub(r'[\\/:*?"<>|]', '_', data[key_field])
                        filename = f"{safe_filename}.docx"
                    else:
                        filename = f"文档_{i + 1}.docx"
                    
                    output_path = os.path.join(output_dir, filename)
                    
                    counter = 1
                    while os.path.exists(output_path):
                        name, ext = os.path.splitext(filename)
                        output_path = os.path.join(output_dir, f"{name}_{counter}{ext}")
                        counter += 1
                    
                    generator.generate_document(data, output_path)
                    results.append(output_path)
                    success_count += 1
                
                except Exception as e:
                    error_count += 1
                    print(f"生成失败: {str(e)}")
                
                self.progress_bar.setValue(i + 1)
            
            operation_id = self.db_manager.add_operation(
                operation_type="批量生成文档",
                template_name=os.path.basename(self.template_path),
                output_path=output_dir,
                record_count=len(data_list),
                success_count=success_count,
                error_count=error_count,
                status="completed" if error_count == 0 else "partially_failed"
            )
            
            for result in results:
                try:
                    file_size = os.path.getsize(result)
                except:
                    file_size = 0
                self.db_manager.add_generated_file(
                    operation_id=operation_id,
                    file_name=os.path.basename(result),
                    file_path=result,
                    file_size=file_size
                )
            
            self.status_label.setText(f"完成: 成功{success_count}个, 失败{error_count}个")
            
            if error_count == 0:
                QMessageBox.information(
                    self, "成功",
                    f"成功生成 {success_count} 个文档\n\n"
                    f"输出目录: {output_dir}"
                )
            else:
                QMessageBox.warning(
                    self, "完成",
                    f"生成完成\n成功: {success_count} 个\n失败: {error_count} 个"
                )
        
        except Exception as e:
            self.status_label.setText("出错")
            QMessageBox.critical(self, "错误", f"生成失败: {str(e)}")
        
        finally:
            self.generate_btn.setEnabled(True)


class FormatTab(QWidget):
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.selected_files = []
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        files_group = QGroupBox("1. 选择要排版的文档")
        files_layout = QVBoxLayout(files_group)
        
        select_btn_layout = QHBoxLayout()
        self.select_files_btn = QPushButton("选择文件...")
        self.select_files_btn.clicked.connect(self.select_files)
        
        self.select_folder_btn = QPushButton("选择文件夹...")
        self.select_folder_btn.clicked.connect(self.select_folder)
        
        self.clear_btn = QPushButton("清空列表")
        self.clear_btn.clicked.connect(self.clear_files)
        
        select_btn_layout.addWidget(self.select_files_btn)
        select_btn_layout.addWidget(self.select_folder_btn)
        select_btn_layout.addWidget(self.clear_btn)
        select_btn_layout.addStretch()
        
        files_layout.addLayout(select_btn_layout)
        
        self.files_table = QTableWidget()
        self.files_table.setColumnCount(2)
        self.files_table.setHorizontalHeaderLabels(["文件名", "路径"])
        self.files_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.files_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        
        files_layout.addWidget(self.files_table)
        
        layout.addWidget(files_group)
        
        format_group = QGroupBox("2. 排版设置")
        format_layout = QVBoxLayout(format_group)
        
        font_layout = QHBoxLayout()
        
        self.font_name_combo = QComboBox()
        self.font_name_combo.addItems(["宋体", "微软雅黑", "黑体", "楷体", "Times New Roman", "Arial"])
        self.font_name_combo.setCurrentText("宋体")
        
        self.font_size_spin = QDoubleSpinBox()
        self.font_size_spin.setRange(1, 72)
        self.font_size_spin.setValue(12)
        self.font_size_spin.setSingleStep(0.5)
        
        self.bold_check = QCheckBox("加粗")
        self.italic_check = QCheckBox("斜体")
        
        font_layout.addWidget(QLabel("字体:"))
        font_layout.addWidget(self.font_name_combo)
        font_layout.addWidget(QLabel("字号:"))
        font_layout.addWidget(self.font_size_spin)
        font_layout.addWidget(self.bold_check)
        font_layout.addWidget(self.italic_check)
        font_layout.addStretch()
        
        format_layout.addLayout(font_layout)
        
        paragraph_layout = QHBoxLayout()
        
        self.line_spacing_spin = QDoubleSpinBox()
        self.line_spacing_spin.setRange(0.5, 5.0)
        self.line_spacing_spin.setValue(1.5)
        self.line_spacing_spin.setSingleStep(0.1)
        
        self.indent_spin = QDoubleSpinBox()
        self.indent_spin.setRange(0, 10)
        self.indent_spin.setValue(2.0)
        self.indent_spin.setSingleStep(0.5)
        self.indent_spin.setSuffix(" 字符")
        
        self.alignment_combo = QComboBox()
        self.alignment_combo.addItems(["左对齐", "居中对齐", "右对齐", "两端对齐"])
        
        paragraph_layout.addWidget(QLabel("行距:"))
        paragraph_layout.addWidget(self.line_spacing_spin)
        paragraph_layout.addWidget(QLabel("首行缩进:"))
        paragraph_layout.addWidget(self.indent_spin)
        paragraph_layout.addWidget(QLabel("对齐方式:"))
        paragraph_layout.addWidget(self.alignment_combo)
        paragraph_layout.addStretch()
        
        format_layout.addLayout(paragraph_layout)
        
        margin_group = QGroupBox("页面边距 (厘米)")
        margin_layout = QHBoxLayout(margin_group)
        
        self.margin_top_spin = QDoubleSpinBox()
        self.margin_top_spin.setRange(0, 10)
        self.margin_top_spin.setValue(2.54)
        
        self.margin_bottom_spin = QDoubleSpinBox()
        self.margin_bottom_spin.setRange(0, 10)
        self.margin_bottom_spin.setValue(2.54)
        
        self.margin_left_spin = QDoubleSpinBox()
        self.margin_left_spin.setRange(0, 10)
        self.margin_left_spin.setValue(3.17)
        
        self.margin_right_spin = QDoubleSpinBox()
        self.margin_right_spin.setRange(0, 10)
        self.margin_right_spin.setValue(3.17)
        
        margin_layout.addWidget(QLabel("上:"))
        margin_layout.addWidget(self.margin_top_spin)
        margin_layout.addWidget(QLabel("下:"))
        margin_layout.addWidget(self.margin_bottom_spin)
        margin_layout.addWidget(QLabel("左:"))
        margin_layout.addWidget(self.margin_left_spin)
        margin_layout.addWidget(QLabel("右:"))
        margin_layout.addWidget(self.margin_right_spin)
        margin_layout.addStretch()
        
        format_layout.addWidget(margin_group)
        
        layout.addWidget(format_group)
        
        output_group = QGroupBox("3. 输出设置")
        output_layout = QHBoxLayout(output_group)
        
        self.output_edit = QLineEdit(OUTPUT_DIR)
        self.output_btn = QPushButton("浏览...")
        self.output_btn.clicked.connect(self.select_output_dir)
        
        output_layout.addWidget(QLabel("输出目录:"))
        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(self.output_btn)
        
        layout.addWidget(output_group)
        
        progress_group = QGroupBox("操作进度")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #666;")
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        
        layout.addWidget(progress_group)
        
        btn_layout = QHBoxLayout()
        
        self.format_btn = QPushButton("开始排版")
        self.format_btn.clicked.connect(self.start_formatting)
        self.format_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.format_btn)
        
        layout.addLayout(btn_layout)
    
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Word文档", "",
            "Word文档 (*.docx);;所有文件 (*)"
        )
        if files:
            self.selected_files.extend(files)
            self.update_files_table()
    
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            docx_files = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.docx'):
                        docx_files.append(os.path.join(root, file))
            
            if docx_files:
                self.selected_files.extend(docx_files)
                self.update_files_table()
                QMessageBox.information(
                    self, "提示",
                    f"从文件夹中找到 {len(docx_files)} 个Word文档"
                )
            else:
                QMessageBox.information(self, "提示", "文件夹中没有找到Word文档")
    
    def clear_files(self):
        self.selected_files = []
        self.update_files_table()
    
    def update_files_table(self):
        self.files_table.setRowCount(len(self.selected_files))
        for i, file_path in enumerate(self.selected_files):
            self.files_table.setItem(i, 0, QTableWidgetItem(os.path.basename(file_path)))
            self.files_table.setItem(i, 1, QTableWidgetItem(file_path))
    
    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_edit.text())
        if dir_path:
            self.output_edit.setText(dir_path)
    
    def get_alignment_value(self) -> str:
        alignment_map = {
            "左对齐": "left",
            "居中对齐": "center",
            "右对齐": "right",
            "两端对齐": "justify"
        }
        return alignment_map.get(self.alignment_combo.currentText(), "left")
    
    def start_formatting(self):
        if not self.selected_files:
            QMessageBox.warning(self, "提示", "请先选择要排版的文档")
            return
        
        output_dir = self.output_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "提示", "请选择输出目录")
            return
        
        font_name = self.font_name_combo.currentText()
        font_size = self.font_size_spin.value()
        bold = self.bold_check.isChecked() if self.bold_check.isChecked() else None
        italic = self.italic_check.isChecked() if self.italic_check.isChecked() else None
        line_spacing = self.line_spacing_spin.value()
        indentation = self.indent_spin.value()
        alignment = self.get_alignment_value()
        
        margins = {
            'top': self.margin_top_spin.value(),
            'bottom': self.margin_bottom_spin.value(),
            'left': self.margin_left_spin.value(),
            'right': self.margin_right_spin.value()
        }
        
        self.progress_bar.setMaximum(len(self.selected_files))
        self.progress_bar.setValue(0)
        self.status_label.setText("正在排版...")
        self.format_btn.setEnabled(False)
        
        try:
            success_count = 0
            error_count = 0
            
            for i, file_path in enumerate(self.selected_files):
                try:
                    from document_generator import DocumentFormatter
                    
                    formatter = DocumentFormatter(file_path)
                    formatter.apply_formatting(
                        font_name=font_name,
                        font_size=font_size,
                        bold=bold,
                        italic=italic,
                        line_spacing=line_spacing,
                        indentation=indentation,
                        alignment=alignment,
                        margins=margins
                    )
                    
                    filename = os.path.basename(file_path)
                    output_path = os.path.join(output_dir, filename)
                    
                    counter = 1
                    while os.path.exists(output_path):
                        name, ext = os.path.splitext(filename)
                        output_path = os.path.join(output_dir, f"{name}_formatted{ext}")
                        counter += 1
                    
                    formatter.save(output_path)
                    success_count += 1
                
                except Exception as e:
                    error_count += 1
                    print(f"排版失败: {str(e)}")
                
                self.progress_bar.setValue(i + 1)
            
            self.db_manager.add_format_operation(
                operation_type="批量排版",
                files_count=len(self.selected_files),
                font_name=font_name,
                font_size=font_size,
                line_spacing=line_spacing,
                indentation=indentation,
                bold=str(bold) if bold is not None else None,
                italic=str(italic) if italic is not None else None,
                status="completed" if error_count == 0 else "partially_failed"
            )
            
            self.status_label.setText(f"完成: 成功{success_count}个, 失败{error_count}个")
            
            if error_count == 0:
                QMessageBox.information(
                    self, "成功",
                    f"成功排版 {success_count} 个文档\n\n"
                    f"输出目录: {output_dir}"
                )
            else:
                QMessageBox.warning(
                    self, "完成",
                    f"排版完成\n成功: {success_count} 个\n失败: {error_count} 个"
                )
        
        except Exception as e:
            self.status_label.setText("出错")
            QMessageBox.critical(self, "错误", f"排版失败: {str(e)}")
        
        finally:
            self.format_btn.setEnabled(True)


class HistoryTab(QWidget):
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.chart_generator = ChartGenerator(db_manager)
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        splitter = QSplitter(Qt.Vertical)
        
        chart_group = QGroupBox("统计图表")
        chart_layout = QVBoxLayout(chart_group)
        
        chart_control_layout = QHBoxLayout()
        
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems([
            "操作统计概览",
            "每日操作统计",
            "操作类型分布",
            "操作成功率",
            "每周操作分布"
        ])
        
        self.days_combo = QComboBox()
        self.days_combo.addItems(["最近7天", "最近30天", "最近90天"])
        self.days_combo.setCurrentIndex(1)
        
        self.refresh_chart_btn = QPushButton("刷新图表")
        self.refresh_chart_btn.clicked.connect(self.refresh_chart)
        
        chart_control_layout.addWidget(QLabel("图表类型:"))
        chart_control_layout.addWidget(self.chart_type_combo)
        chart_control_layout.addWidget(QLabel("时间范围:"))
        chart_control_layout.addWidget(self.days_combo)
        chart_control_layout.addWidget(self.refresh_chart_btn)
        chart_control_layout.addStretch()
        
        chart_layout.addLayout(chart_control_layout)
        
        self.chart_label = QLabel()
        self.chart_label.setAlignment(Qt.AlignCenter)
        self.chart_label.setMinimumHeight(350)
        self.chart_label.setStyleSheet("background-color: white; border: 1px solid #ccc;")
        
        chart_layout.addWidget(self.chart_label)
        
        splitter.addWidget(chart_group)
        
        history_group = QGroupBox("操作记录")
        history_layout = QVBoxLayout(history_group)
        
        history_control_layout = QHBoxLayout()
        
        self.history_type_combo = QComboBox()
        self.history_type_combo.addItems(["所有操作", "文档生成", "排版操作"])
        
        self.refresh_history_btn = QPushButton("刷新记录")
        self.refresh_history_btn.clicked.connect(self.refresh_history)
        
        history_control_layout.addWidget(QLabel("操作类型:"))
        history_control_layout.addWidget(self.history_type_combo)
        history_control_layout.addWidget(self.refresh_history_btn)
        history_control_layout.addStretch()
        
        history_layout.addLayout(history_control_layout)
        
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(7)
        self.history_table.setHorizontalHeaderLabels([
            "ID", "操作类型", "模板/文件数", "成功", "失败", "时间", "状态"
        ])
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        history_layout.addWidget(self.history_table)
        
        splitter.addWidget(history_group)
        
        layout.addWidget(splitter)
        
        self.refresh_chart()
        self.refresh_history()
    
    def get_days_value(self) -> int:
        days_map = {
            "最近7天": 7,
            "最近30天": 30,
            "最近90天": 90
        }
        return days_map.get(self.days_combo.currentText(), 30)
    
    def refresh_chart(self):
        try:
            chart_type = self.chart_type_combo.currentText()
            days = self.get_days_value()
            
            if chart_type == "操作统计概览":
                fig = self.chart_generator.create_operation_statistics_chart(days)
            elif chart_type == "每日操作统计":
                fig = self.chart_generator.create_daily_operations_chart(days)
            elif chart_type == "操作类型分布":
                fig = self.chart_generator.create_operation_type_pie_chart()
            elif chart_type == "操作成功率":
                fig = self.chart_generator.create_success_rate_chart()
            elif chart_type == "每周操作分布":
                fig = self.chart_generator.create_weekly_operations_chart()
            else:
                fig = self.chart_generator.create_operation_statistics_chart(days)
            
            img_bytes = self.chart_generator.figure_to_bytes(fig)
            
            pixmap = QPixmap()
            pixmap.loadFromData(img_bytes)
            
            label_size = self.chart_label.size()
            scaled_pixmap = pixmap.scaled(
                label_size.width() - 20,
                label_size.height() - 20,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            
            self.chart_label.setPixmap(scaled_pixmap)
        
        except Exception as e:
            self.chart_label.setText(f"无法显示图表: {str(e)}")
    
    def refresh_history(self):
        try:
            history_type = self.history_type_combo.currentText()
            
            if history_type == "所有操作":
                operations = self.db_manager.get_operations(limit=100)
                format_ops = self.db_manager.get_format_operations(limit=50)
            elif history_type == "文档生成":
                operations = self.db_manager.get_operations_by_type("批量生成文档", limit=100)
                format_ops = []
            else:
                operations = []
                format_ops = self.db_manager.get_format_operations(limit=100)
            
            all_records = []
            
            for op in operations:
                all_records.append({
                    'id': op['id'],
                    'type': '文档生成',
                    'template': op.get('template_name', '-'),
                    'success': op.get('success_count', 0),
                    'error': op.get('error_count', 0),
                    'time': op.get('created_at', '-'),
                    'status': op.get('status', 'completed')
                })
            
            for op in format_ops:
                all_records.append({
                    'id': op['id'],
                    'type': '排版操作',
                    'template': f"{op.get('files_count', 0)} 个文件",
                    'success': op.get('files_count', 0) if op.get('status') == 'completed' else 0,
                    'error': 0,
                    'time': op.get('created_at', '-'),
                    'status': op.get('status', 'completed')
                })
            
            all_records.sort(key=lambda x: x['time'], reverse=True)
            
            self.history_table.setRowCount(len(all_records))
            
            status_color_map = {
                'completed': '#2ecc71',
                'partially_failed': '#f39c12',
                'failed': '#e74c3c'
            }
            
            for i, record in enumerate(all_records):
                self.history_table.setItem(i, 0, QTableWidgetItem(str(record['id'])))
                self.history_table.setItem(i, 1, QTableWidgetItem(record['type']))
                self.history_table.setItem(i, 2, QTableWidgetItem(record['template']))
                self.history_table.setItem(i, 3, QTableWidgetItem(str(record['success'])))
                self.history_table.setItem(i, 4, QTableWidgetItem(str(record['error'])))
                self.history_table.setItem(i, 5, QTableWidgetItem(str(record['time'])))
                
                status_item = QTableWidgetItem(
                    "完成" if record['status'] == 'completed' 
                    else "部分完成" if record['status'] == 'partially_failed'
                    else "失败"
                )
                color = status_color_map.get(record['status'], '#666')
                self.history_table.setItem(i, 6, status_item)
        
        except Exception as e:
            print(f"刷新记录失败: {str(e)}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db_manager = DatabaseManager()
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(1000, 750)
        self.resize(1200, 800)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        header_label = QLabel(APP_NAME)
        header_label.setFont(QFont("Microsoft YaHei", 20, QFont.Bold))
        header_label.setAlignment(Qt.AlignCenter)
        header_label.setStyleSheet("color: #2c3e50; padding: 10px;")
        
        layout.addWidget(header_label)
        
        self.tab_widget = QTabWidget()
        self.tab_widget.setFont(QFont("Microsoft YaHei", 10))
        
        self.generate_tab = GenerateTab(self.db_manager)
        self.format_tab = FormatTab(self.db_manager)
        self.history_tab = HistoryTab(self.db_manager)
        
        self.tab_widget.addTab(self.generate_tab, "📄 批量生成文档")
        self.tab_widget.addTab(self.format_tab, "🎨 批量排版")
        self.tab_widget.addTab(self.history_tab, "📊 使用记录")
        
        layout.addWidget(self.tab_widget)
        
        status_label = QLabel(f"{APP_NAME} v{APP_VERSION} - 准备就绪")
        status_label.setStyleSheet("color: #7f8c8d; padding: 5px;")
        self.statusBar().addWidget(status_label)


def main():
    app = QApplication(sys.argv)
    
    app.setStyle('Fusion')
    
    from PyQt5.QtGui import QPalette, QColor
    
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.WindowText, QColor(44, 62, 80))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(245, 246, 250))
    palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipText, QColor(44, 62, 80))
    palette.setColor(QPalette.Text, QColor(44, 62, 80))
    palette.setColor(QPalette.Button, QColor(245, 246, 250))
    palette.setColor(QPalette.ButtonText, QColor(44, 62, 80))
    palette.setColor(QPalette.BrightText, QColor(231, 76, 60))
    palette.setColor(QPalette.Highlight, QColor(52, 152, 219))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    
    app.setPalette(palette)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
