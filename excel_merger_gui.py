import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog,
                            QListWidget, QLineEdit, QMessageBox, QProgressBar,
                            QFrame)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette
from merge_excel import merge_excel_files

class ExcelMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.output_file_path = None
        
    def initUI(self):
        self.setWindowTitle('Excel 合并工具')
        self.setMinimumSize(600, 400)
        
        # 设置iOS风格配色
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(242, 242, 247))
        palette.setColor(QPalette.WindowText, QColor(28, 28, 30))
        palette.setColor(QPalette.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.AlternateBase, QColor(242, 242, 247))
        palette.setColor(QPalette.Button, QColor(0, 122, 255))
        palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))
        self.setPalette(palette)
        
        # 设置字体
        font = QFont("SF Pro Text", 13)
        self.setFont(font)
        
        # 主窗口布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 添加圆角效果
        main_widget.setStyleSheet("""
            QWidget {
                background-color: #F2F2F7;
                border-radius: 12px;
            }
            QPushButton {
                background-color: #5AC8FA;
                color: white;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #4AB2E2;
            }
            QLineEdit, QListWidget {
                border: 1px solid #C6C6C8;
                border-radius: 8px;
                padding: 8px;
                background-color: white;
            }
            QLabel {
                color: #1C1C1E;
                font-size: 14px;
            }
        """)
        
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
        
        # 文件名输入和打开按钮
        file_name_layout = QHBoxLayout()
        self.file_name = QLineEdit()
        self.file_name.setPlaceholderText("输入输出文件名（不带扩展名）")
        file_name_layout.addWidget(self.file_name)
        
        # 打开文件按钮
        self.open_button = QPushButton("打开文件")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self.open_output_file)
        file_name_layout.addWidget(self.open_button)
        layout.addLayout(file_name_layout)
        
        # 进度条
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)
        
        # 状态标签
        self.status_label = QLabel()
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
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
        """选择输出路径并显示Excel文件"""
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_path.setText(path)
            # 启用打开按钮如果存在Excel文件
            output_file = os.path.join(path, f"{self.file_name.text()}.xlsx")
            if os.path.exists(output_file):
                self.output_file_path = output_file
                self.open_button.setEnabled(True)
            else:
                self.open_button.setEnabled(False)
                
            # 显示目录下的Excel文件
            excel_files = [f for f in os.listdir(path) 
                         if f.lower().endswith(('.xls', '.xlsx'))]
            if excel_files:
                self.file_list.clear()
                self.file_list.addItems([os.path.join(path, f) for f in excel_files])
                
    def open_output_file(self):
        """打开生成的Excel文件"""
        if self.output_file_path and os.path.exists(self.output_file_path):
            try:
                os.startfile(self.output_file_path)
            except Exception as e:
                QMessageBox.warning(self, "错误", f"无法打开文件：\n{str(e)}")
        else:
            QMessageBox.warning(self, "错误", "文件不存在或路径无效")
            
    def merge_files(self):
        # 验证输入
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
        self.output_file_path = os.path.join(self.output_path.text(), 
                                 f"{self.file_name.text()}.xlsx")
        
        # 检查输出文件是否已存在
        if os.path.exists(self.output_file_path):
            reply = QMessageBox.question(self, '文件已存在',
                                       '输出文件已存在，是否覆盖？',
                                       QMessageBox.Yes | QMessageBox.No,
                                       QMessageBox.No)
            if reply == QMessageBox.No:
                return
        
        # 初始化进度条
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status_label.setText("正在合并文件...")
        
        # 使用定时器模拟异步操作
        QTimer.singleShot(100, lambda: self.do_merge(file_paths, self.output_file_path))
            
    def do_merge(self, file_paths, output_file):
        try:
            # 执行合并
            merge_excel_files(file_paths, output_file)
            
            # 更新进度
            self.progress.setValue(100)
            self.status_label.setText("合并完成")
            
            # 显示成功消息并启用打开按钮
            QMessageBox.information(self, "成功", 
                                  f"文件已成功合并到：\n{output_file}")
            self.open_button.setEnabled(True)
            
        except Exception as e:
            # 删除可能创建的部分文件
            if os.path.exists(output_file):
                try:
                    os.remove(output_file)
                except:
                    pass
                    
            # 显示错误信息
            self.progress.setValue(0)
            self.status_label.setText("合并失败")
            QMessageBox.critical(self, "错误", 
                               f"合并过程中发生错误：\n{str(e)}")
            self.open_button.setEnabled(False)
            
        finally:
            # 重置进度条
            QTimer.singleShot(2000, lambda: self.progress.setVisible(False))
            
    def clear_list(self):
        self.file_list.clear()
        self.status_label.clear()
        self.open_button.setEnabled(False)
        
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
