import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog,
                            QListWidget, QLineEdit, QMessageBox)
from PyQt5.QtCore import Qt
from merge_excel import merge_excel_files

class ExcelMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Excel 合并工具')
        self.setMinimumSize(600, 400)  # 设置最小尺寸
        
        # 主窗口布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)  # 减少边距
        layout.setSpacing(10)  # 减少控件间距
        
        # 文件列表区域
        file_select_layout = QHBoxLayout()
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        
        select_button = QPushButton("选择文件...")
        select_button.clicked.connect(self.select_input_files)
        
        file_select_layout.addWidget(QLabel("选择Excel文件："))
        file_select_layout.addWidget(select_button)
        layout.addLayout(file_select_layout)
        layout.addWidget(self.file_list)
        
        # 输出路径选择
        output_layout = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("选择输出路径")
        output_layout.addWidget(self.output_path)
        
        browse_button = QPushButton("浏览...")
        browse_button.clicked.connect(self.select_output_path)
        output_layout.addWidget(browse_button)
        layout.addLayout(output_layout)
        
        # 文件名输入
        self.file_name = QLineEdit()
        self.file_name.setPlaceholderText("输入输出文件名（不带扩展名）")
        layout.addWidget(self.file_name)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        merge_button = QPushButton("合并文件")
        merge_button.clicked.connect(self.merge_files)
        button_layout.addWidget(merge_button)
        
        clear_button = QPushButton("清空列表")
        clear_button.clicked.connect(self.clear_list)
        button_layout.addWidget(clear_button)
        layout.addLayout(button_layout)
        
        main_widget.setLayout(layout)
        
    def select_output_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_path.setText(path)
            
    def merge_files(self):
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "错误", "请先添加要合并的文件")
            return
            
        if not self.output_path.text():
            QMessageBox.warning(self, "错误", "请选择输出路径")
            return
            
        if not self.file_name.text():
            QMessageBox.warning(self, "错误", "请输入输出文件名")
            return
            
        # 获取文件列表
        file_paths = [self.file_list.item(i).text() 
                     for i in range(self.file_list.count())]
        
        # 构建完整输出路径
        output_file = os.path.join(self.output_path.text(), 
                                 f"{self.file_name.text()}.xlsx")
        
        try:
            merge_excel_files(file_paths, output_file)
            QMessageBox.information(self, "成功", 
                                  f"文件已成功合并到：\n{output_file}")
        except Exception as e:
            QMessageBox.critical(self, "错误", 
                               f"合并过程中发生错误：\n{str(e)}")
            
    def clear_list(self):
        self.file_list.clear()
        
    def select_input_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择Excel文件",
            "",
            "Excel 文件 (*.xls *.xlsx)"
        )
        
        if files:
            for file_path in files:
                # 检查文件扩展名
                if file_path.lower().endswith(('.xls', '.xlsx')):
                    # 检查文件是否存在
                    if os.path.exists(file_path):
                        # 检查是否已经添加过
                        if not self.file_list.findItems(file_path, Qt.MatchExactly):
                            self.file_list.addItem(file_path)
                        else:
                            QMessageBox.warning(self, "提示", f"文件已存在: {file_path}")
                    else:
                        QMessageBox.warning(self, "错误", f"文件不存在: {file_path}")
                else:
                    QMessageBox.warning(self, "错误", f"文件格式不支持: {file_path}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelMergerApp()
    window.show()
    sys.exit(app.exec_())
