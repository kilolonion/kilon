import sys
import os
import pandas as pd
import numpy as np
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QListWidget, QTableWidget, QTableWidgetItem, QPushButton, 
                             QLabel, QComboBox, QProgressBar, QTextEdit, QFileDialog, 
                             QMessageBox, QInputDialog, QDialog, QCheckBox, QListWidgetItem, 
                             QScrollArea, QLineEdit, QDialogButtonBox, QMenu, QStyle, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QMutex
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon
import logging
from datetime import datetime
import configparser
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

from help_dialog import HelpDialog
from settings_dialog import SettingsDialog
from updater import Updater, show_update_dialog, show_update_completed_dialog
from welcome_dialog import WelcomeDialog

class AdvancedVLOOKUPTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("高级VLOOKUP工具")
        self.setGeometry(100, 100, 1200, 800)

        icon_path = os.path.join(os.path.dirname(__file__), 'VLookUp.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))  # 设置应用图标

        self.current_version = self.load_version()  # 从 version.ini 读取当前版本

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)

        self.file_load_progress = QProgressBar()
        self.file_load_progress.setVisible(False)  # 初始化文件加载进度条

        self.recent_files = []
        self.settings = QSettings("YourCompany", "AdvancedVLOOKUPTool")
        self.config = self.load_config()  # 初始化最近使用的文件列表和设置
        
        self.setup_ui()
        self.setup_menu()
        self.load_settings()

        self.loaded_files = {}
        self.main_table = None
        self.lookup_tables = []  # 初始化文件和表格相关变量

        self.setAcceptDrops(True)  # 允许拖放操作

        self.ui_mutex = QMutex()  # 初始化UI互斥锁

        logging.basicConfig(filename='vlookup_tool.log', level=logging.INFO, 
                            format='%(asctime)s - %(levelname)s - %(message)s')  # 设置日志

        self.auto_update_check = self.settings.value("auto_update_check", True, type=bool)
        if self.auto_update_check:
            self.check_for_updates()  # 检查自动更新设置

        if not self.settings.value("hide_welcome", False, type=bool):
            welcome = WelcomeDialog(self)
            if welcome.exec() == QDialog.DialogCode.Accepted and welcome.dont_show_again.isChecked():
                self.settings.setValue("hide_welcome", True)  # 显示欢迎对话框

        self.update_url = "https://api.github.com/repos/yourusername/AdvancedVLOOKUPTool/releases/latest"
        self.updater = Updater(self.current_version, self.update_url)
        self.setup_updater()  # 设置更新器

    def setup_updater(self):
        self.updater.update_available.connect(self.on_update_available)
        self.updater.update_progress.connect(self.on_update_progress)
        self.updater.update_completed.connect(self.on_update_completed)
        self.updater.update_error.connect(self.on_update_error)  # 设置更新器的信号连接

    def on_update_available(self, version):
        if show_update_dialog(self, version):
            self.updater.update_application(version)  # 当有新版本可用时的处理函数

    def on_update_progress(self, percent):
        try:
            percent = max(0, min(100, percent))  # 确保百分比在有效范围内
            
            if hasattr(self, 'progress_bar'):
                self.progress_bar.setValue(percent)  # 更新进度条
            
            if hasattr(self, 'statusBar'):
                self.statusBar().showMessage(f"更新进度: {percent}%")  # 更新状态栏消息
            
            print(f"更新进度: {percent}%")  # 如果需要，可以在控制台输出进度
            
            if percent == 100:
                QMessageBox.information(self, "更新完成", "软件更新已完成！")  # 如果进度完成，可以显示一个消息
                
            QApplication.processEvents()  # 确保 GUI 事件得到处理
            
        except Exception as e:
            print(f"更新进度时发生错误: {str(e)}")
            logging.error(f"更新进度时发生错误: {str(e)}")  # 可以选择记录错误到日志文件

    def on_update_completed(self):
        show_update_completed_dialog(self)
        self.close()  # 更新完成后的处理函数

    def on_update_error(self, error_message):
        QMessageBox.warning(self, "更新错误", error_message)  # 更新错误时的处理函数

    def load_config(self):
        config = configparser.ConfigParser()
        config_file = 'config.ini'
        if os.path.exists(config_file):
            config.read(config_file)
        else:
            config['DEFAULT'] = {
                'ChunkSize': '100000',
                'MaxRecentFiles': '5',
                'DefaultSaveFormat': 'xlsx'
            }
            with open(config_file, 'w') as configfile:
                config.write(configfile)  # 如果配置文件不存在，创建默认配置
        return config

    def setup_ui(self):
        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        # 文件列表和按钮
        file_list_layout = QHBoxLayout()
        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self.preview_file)
        file_list_layout.addWidget(self.file_list)

        file_buttons_layout = QVBoxLayout()
        self.select_file_button = QPushButton("选择文件")
        self.select_file_button.clicked.connect(self.load_files)
        self.recent_files_button = QPushButton("最近使用的文件")
        self.recent_files_button.clicked.connect(self.show_recent_files)
        self.delete_file_button = QPushButton("删除文件")
        self.delete_file_button.clicked.connect(self.delete_selected_file)
        file_buttons_layout.addWidget(self.select_file_button)
        file_buttons_layout.addWidget(self.recent_files_button)
        file_buttons_layout.addWidget(self.delete_file_button)
        file_list_layout.addLayout(file_buttons_layout)

        left_layout.addWidget(QLabel("已加载文件:"))
        left_layout.addLayout(file_list_layout)

        # 预览表格
        self.preview_table = QTableWidget()
        left_layout.addWidget(QLabel("文件预览:"))
        left_layout.addWidget(self.preview_table)

        # 主表选择
        self.main_table_combo = QComboBox()
        self.main_table_combo.currentIndexChanged.connect(self.update_main_column_combo)
        right_layout.addWidget(QLabel("选择主表:"))
        right_layout.addWidget(self.main_table_combo)

        # 主列选择
        self.main_column_combo = QComboBox()
        right_layout.addWidget(QLabel("选择主列:"))
        right_layout.addWidget(self.main_column_combo)

        # 查找表选择
        self.lookup_table_list = QListWidget()
        self.lookup_table_list.itemChanged.connect(self.update_lookup_column_combos)
        right_layout.addWidget(QLabel("选择查找表:"))
        right_layout.addWidget(self.lookup_table_list)

        # 查找列选择
        self.lookup_column_combos = {}
        self.lookup_column_widget = QWidget()
        self.lookup_column_layout = QVBoxLayout(self.lookup_column_widget)
        right_layout.addWidget(QLabel("选择查找列:"))
        right_layout.addWidget(self.lookup_column_widget)

        # 返回列选择
        self.return_columns_list = QListWidget()
        right_layout.addWidget(QLabel("选择返回列:"))
        right_layout.addWidget(self.return_columns_list)

        # 返回列过滤
        self.return_columns_filter = QLineEdit()
        self.return_columns_filter.setPlaceholderText("过滤返回列...")
        self.return_columns_filter.textChanged.connect(self.filter_return_columns)
        right_layout.addWidget(self.return_columns_filter)

        # 执行按钮
        self.execute_button = QPushButton("执行VLOOKUP")
        self.execute_button.clicked.connect(self.execute_vlookup)
        right_layout.addWidget(self.execute_button)

        # 进度条
        self.progress_bar = QProgressBar()
        right_layout.addWidget(self.progress_bar)

        # 结果表格
        self.result_table = QTableWidget()
        right_layout.addWidget(QLabel("VLOOKUP结果:"))
        right_layout.addWidget(self.result_table)

        # 保存结果按钮
        self.save_button = QPushButton("保存结果")
        self.save_button.clicked.connect(self.save_results)
        right_layout.addWidget(self.save_button)

        # 日志文本框
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        right_layout.addWidget(QLabel("操作日志:"))
        right_layout.addWidget(self.log_text)

        # 添加表头选择布局
        self.table_header_widget = QWidget()
        self.table_header_layout = QVBoxLayout(self.table_header_widget)
        right_layout.addWidget(QLabel("表头选择:"))
        right_layout.addWidget(self.table_header_widget)

        # 在右侧布局的底部添加版本号
        version_layout = QHBoxLayout()
        version_layout.addWidget(QLabel("当前版本:"))
        version_layout.addWidget(QLabel(self.current_version))
        right_layout.addLayout(version_layout)

        # 设置主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)
        self.setCentralWidget(main_widget)

    def setup_menu(self):
        menubar = self.menuBar()

        # 文件菜单
        self.file_menu = menubar.addMenu('文件')
        load_action = self.file_menu.addAction('加载文件')
        load_action.triggered.connect(self.load_files)
        
        self.recent_files_menu = self.file_menu.addMenu('最近使用的文件')
        
        clear_action = self.file_menu.addAction('清除所有文件')
        clear_action.triggered.connect(self.clear_files)
        
        self.file_menu.addSeparator()
        exit_action = self.file_menu.addAction('退出')
        exit_action.triggered.connect(self.close)

        # 工具菜单
        tools_menu = menubar.addMenu('工具')
        clean_data_action = tools_menu.addAction('数据清理')
        clean_data_action.triggered.connect(self.clean_data)

        # 设置菜单
        settings_menu = menubar.addMenu('设置')
        settings_action = settings_menu.addAction('首选项')
        settings_action.triggered.connect(self.show_settings)

        # 帮助菜单
        help_menu = menubar.addMenu('帮助')
        help_action = help_menu.addAction('查看帮助')
        help_action.triggered.connect(self.show_help)
        about_action = help_menu.addAction('关于')
        about_action.triggered.connect(self.show_about)

    def load_settings(self):
        self.auto_update_check = self.settings.value("auto_update_check", True, type=bool)
        self.default_save_format = self.settings.value("default_save_format", "xlsx")
        self.max_recent_files = int(self.config.get('DEFAULT', 'MaxRecentFiles', fallback=5))
        
        recent_files = self.settings.value("recent_files", [])
        self.recent_files = []
        for item in recent_files:
            if isinstance(item, (list, tuple)) and len(item) == 2:
                path, time_str = item
                if os.path.exists(path):
                    try:
                        time = datetime.fromisoformat(time_str)
                    except (ValueError, TypeError):
                        time = datetime.now()  # 如果时间格式无效，使用当前时间
                    self.recent_files.append((path, time))
            elif isinstance(item, str) and os.path.exists(item):
                self.recent_files.append((item, datetime.now()))
        
        self.update_recent_files_menu()

    def update_recent_files_menu(self):
        self.recent_files_menu.clear()
        for file_path, _ in self.recent_files:
            action = self.recent_files_menu.addAction(os.path.basename(file_path))
            action.triggered.connect(lambda checked, path=file_path: self.load_file(path))
        
        if self.recent_files:
            self.recent_files_menu.addSeparator()
            self.recent_files_menu.addAction("清除最近使用的文件", self.clear_recent_files)

    def update_recent_files(self, file_path):
        self.recent_files = [(path, time) for path, time in self.recent_files if path != file_path]
        self.recent_files.insert(0, (file_path, datetime.now()))
        self.recent_files = self.recent_files[:self.max_recent_files]
        self.settings.setValue("recent_files", [(path, time.isoformat()) for path, time in self.recent_files])
        self.update_recent_files_menu()

    def load_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        for file_path in file_paths:
            self.load_file(file_path)

    def load_file(self, file_path):
        try:
            file_name = os.path.basename(file_path)
            engine = 'openpyxl' if file_path.endswith('.xlsx') else 'xlrd'
            xl = pd.ExcelFile(file_path, engine=engine)
            sheet_to_df_map = {}
            
            for sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                detected_header = self.detect_header_row(df)
                if detected_header > 0:
                    df.columns = df.iloc[detected_header].astype(str)
                    df = df.drop(df.index[detected_header]).reset_index(drop=True)
                else:
                    df.columns = df.columns.astype(str)
                sheet_to_df_map[sheet_name] = {
                    'data': df,
                    'detected_header': detected_header
                }

            self.loaded_files[file_path] = sheet_to_df_map
            self.file_list.addItem(file_name)
            self.log(f"已加载文件：{file_name}")
            self.update_table_combos()
            self.update_recent_files(file_path)
        except Exception as e:
            self.log(f"加载文件 {file_path} 时发生错误: {str(e)}", logging.ERROR)
            QMessageBox.warning(self, "加载失败", f"文件 {file_name} 加载失败：{str(e)}")

    def delete_selected_file(self):
        current_item = self.file_list.currentItem()
        if current_item is None:
            QMessageBox.warning(self, "警告", "请先选择要删除的文件")
            return
        
        file_name = current_item.text()
        reply = QMessageBox.question(self, '确认删除', 
                                     f"是否确定要删除文件 {file_name}？",
                                     QMessageBox.StandardButton.Yes | 
                                     QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            for file_path in list(self.loaded_files.keys()):
                if os.path.basename(file_path) == file_name:
                    del self.loaded_files[file_path]
                    self.file_list.takeItem(self.file_list.row(current_item))
                    self.log(f"已删除文件：{file_name}")
                    self.update_table_combos()
                    break

    def clear_files(self):
        self.loaded_files.clear()
        self.file_list.clear()
        self.main_table_combo.clear()
        self.return_columns_list.clear()
        self.main_column_combo.clear()
        self.log("已清除所有文件")

    def clear_recent_files(self):
        self.recent_files.clear()
        self.settings.setValue("recent_files", [])
        self.update_recent_files_menu()
        self.log("已清除最近使用的文件列表")

    def show_recent_files(self):
        if not self.recent_files:
            QMessageBox.information(self, "提示", "没有最近使用的文件")
            return

        dialog = RecentFilesDialog(self.recent_files, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            for file_path in dialog.selected_files:
                self.load_file(file_path)

    def update_table_combos(self):
        self.main_table_combo.clear()
        self.lookup_table_list.clear()
        
        # 清除现有的表头选择控件
        if hasattr(self, 'table_header_layout'):
            for i in reversed(range(self.table_header_layout.count())): 
                widget = self.table_header_layout.itemAt(i).widget()
                if widget is not None:
                    widget.setParent(None)
                else:
                    self.table_header_layout.removeItem(self.table_header_layout.itemAt(i))
        else:
            self.log("警告：table_header_layout 不存在", logging.WARNING)
            return
        
        for file_path, sheet_data in self.loaded_files.items():
            file_name = os.path.basename(file_path)
            for sheet_name, sheet_info in sheet_data.items():
                item_text = f"{file_name} - {sheet_name}"
                self.main_table_combo.addItem(item_text)
                item = QListWidgetItem(item_text)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked)
                self.lookup_table_list.addItem(item)

                # 添加表头选择下拉框
                header_combo = QComboBox()
                header_combo.addItem(f"智能检测 (行 {sheet_info['detected_header'] + 1})")
                for i in range(min(10, sheet_info['data'].shape[0])):
                    header_combo.addItem(f"行 {i + 1}")
                header_combo.setCurrentIndex(0)
                header_combo.currentIndexChanged.connect(lambda idx, s=sheet_name, f=file_path: self.update_sheet_header(s, f, idx))
                
                header_layout = QHBoxLayout()
                header_layout.addWidget(QLabel(f"{item_text} 表头:"))
                header_layout.addWidget(header_combo)
                self.table_header_layout.addLayout(header_layout)

        self.update_main_column_combo()
        self.update_lookup_column_combos()
        self.update_return_columns()

    def update_main_column_combo(self):
        self.main_column_combo.clear()
        if self.main_table_combo.currentText():
            df = self.get_dataframe(*self.main_table_combo.currentText().split(" - "))
            if df is not None and isinstance(df, pd.DataFrame):
                # 将所有列名转换为字符串
                df.columns = df.columns.astype(str)
                self.main_column_combo.addItems(df.columns.tolist())
        self.update_lookup_column_combos()

    def update_lookup_column_combos(self):
        for i in reversed(range(self.lookup_column_layout.count())): 
            self.lookup_column_layout.itemAt(i).widget().setParent(None)
        
        self.lookup_column_combos.clear()

        for i in range(self.lookup_table_list.count()):
            item = self.lookup_table_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                file_name, sheet_name = item.text().split(" - ")
                df = self.get_dataframe(file_name, sheet_name)
                if df is not None and isinstance(df, pd.DataFrame):
                    combo = QComboBox()
                    # 将所有列名转换为字符串
                    df.columns = df.columns.astype(str)
                    combo.addItems(df.columns.tolist())
                    main_column = self.main_column_combo.currentText()
                    if main_column in df.columns:
                        combo.setCurrentText(main_column)
                    self.lookup_column_combos[item.text()] = combo
                    self.lookup_column_layout.addWidget(QLabel(f"查找列 ({item.text()}):"))
                    self.lookup_column_layout.addWidget(combo)
        self.update_return_columns()

    def update_return_columns(self):
        self.return_columns_list.clear()
        for i in range(self.lookup_table_list.count()):
            item = self.lookup_table_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                file_name, sheet_name = item.text().split(" - ")
                df = self.get_dataframe(file_name, sheet_name)
                if df is not None and isinstance(df, pd.DataFrame):
                    # 将所有列名转换为字符串
                    df.columns = df.columns.astype(str)
                    for column in df.columns:
                        list_item = QListWidgetItem(self.return_columns_list)
                        checkbox = QCheckBox(str(column))
                        self.return_columns_list.setItemWidget(list_item, checkbox)

    def filter_return_columns(self, text):
        for i in range(self.return_columns_list.count()):
            item = self.return_columns_list.item(i)
            widget = self.return_columns_list.itemWidget(item)
            if isinstance(widget, QCheckBox):
                item.setHidden(text.lower() not in widget.text().lower())

    def get_dataframe(self, file_name, sheet_name):
        for file_path, sheet_data in self.loaded_files.items():
            if os.path.basename(file_path) == file_name:
                return sheet_data[sheet_name]['data']
        return None

    def execute_vlookup(self):
        if not self.validate_vlookup_inputs():
            return

        main_df, main_column, lookup_tables, return_columns = self.get_vlookup_parameters()
        self.vlookup_thread = VLOOKUPThread(main_df, main_column, lookup_tables, return_columns)
        self.vlookup_thread.progress_update.connect(self.progress_bar.setValue)
        self.vlookup_thread.result_ready.connect(self.display_results)
        self.vlookup_thread.error_occurred.connect(self.handle_vlookup_error)
        self.vlookup_thread.start()

    def validate_vlookup_inputs(self):
        if not self.main_table_combo.currentText():
            QMessageBox.warning(self, "警告", "请选择主表")
            return False
        if not any(self.lookup_table_list.item(i).checkState() == Qt.CheckState.Checked for i in range(self.lookup_table_list.count())):
            QMessageBox.warning(self, "警告", "请选择至少一个查找表")
            return False
        if not self.get_selected_return_columns():
            QMessageBox.warning(self, "警告", "请选择至少一个返回列")
            return False
        return True

    def get_vlookup_parameters(self):
        file_name, main_sheet = self.main_table_combo.currentText().split(" - ")
        main_df = self.get_dataframe(file_name, main_sheet)
        main_column = self.main_column_combo.currentText()

        lookup_tables = [
            (self.get_dataframe(*item.text().split(" - ")), self.lookup_column_combos[item.text()].currentText())
            for item in (self.lookup_table_list.item(i) for i in range(self.lookup_table_list.count()))
            if item.checkState() == Qt.CheckState.Checked
        ]

        return_columns = self.get_selected_return_columns()

        return main_df, main_column, lookup_tables, return_columns

    def get_selected_return_columns(self):
        return [self.return_columns_list.itemWidget(self.return_columns_list.item(i)).text() 
                for i in range(self.return_columns_list.count()) 
                if not self.return_columns_list.item(i).isHidden() and 
                self.return_columns_list.itemWidget(self.return_columns_list.item(i)).isChecked()]

    def display_results(self, df):
        self.last_result = df
        self.display_dataframe(df, self.result_table)
        self.log("VLOOKUP执行完成")

    def display_dataframe(self, df, table_widget):
        table_widget.setRowCount(df.shape[0])
        table_widget.setColumnCount(df.shape[1])
        
        # 将列名转换为字符串
        column_labels = [str(col) for col in df.columns]
        table_widget.setHorizontalHeaderLabels(column_labels)

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                table_widget.setItem(row, col, QTableWidgetItem(str(df.iloc[row, col])))

    def handle_vlookup_error(self, error_message):
        self.log(f"VLOOKUP操作错误: {error_message}", logging.ERROR)
        QMessageBox.critical(self, "错误", f"VLOOKUP操作失败: {error_message}")

    def save_results(self):
        if self.last_result is None:
            QMessageBox.warning(self, "警告", "没有可保存的结果")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "保存结果", "", 
                                                   f"{self.default_save_format.upper()} Files (*.{self.default_save_format});;All Files (*)")
        
        if file_path:
            try:
                if file_path.endswith('.xlsx'):
                    self.last_result.to_excel(file_path, index=False)
                elif file_path.endswith('.csv'):
                    self.last_result.to_csv(file_path, index=False)
                elif file_path.endswith('.pdf'):
                    self.save_as_pdf(file_path, self.last_result)
                else:
                    self.last_result.to_excel(file_path + '.xlsx', index=False)
                
                self.log(f"结果已成功保存至：{file_path}")
                QMessageBox.information(self, "保存成功", f"结果已成功保存至：{file_path}")
            except Exception as e:
                self.log(f"保存结果失败：{str(e)}", logging.ERROR)
                QMessageBox.warning(self, "保存失败", f"保存结果时发生错误：{str(e)}")

    def save_as_pdf(self, file_path, df):
        doc = SimpleDocTemplate(file_path, pagesize=landscape(letter))
        elements = []

        styles = getSampleStyleSheet()
        title = Paragraph("VLOOKUP Results", styles['Title'])
        elements.append(title)

        data = [df.columns.tolist()] + df.values.tolist()
        t = Table(data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 14),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(t)

        doc.build(elements)

    def log(self, message, level=logging.INFO):
        logging.log(level, message)
        self.log_text.append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")

    def preview_file(self, item):
        file_name = item.text()
        for file_path, sheet_data in self.loaded_files.items():
            if os.path.basename(file_path) == file_name:
                sheet_name = next(iter(sheet_data.keys()))
                df = sheet_data[sheet_name]['data']  # 获取实际的 DataFrame
                
                # 如果第一行是表头，则设置它为列名
                header_row = sheet_data[sheet_name]['detected_header']
                if header_row is not None:
                    df.columns = df.iloc[header_row]
                    df = df.drop(df.index[header_row])
                
                self.display_dataframe(df.head(10), self.preview_table)
                self.log(f"预览文件：{file_name}, 表：{sheet_name}")
                break

    def clean_data(self):
        if not self.loaded_files:
            QMessageBox.warning(self, "警告", "请先加载文件")
            return

        file_name, ok = QInputDialog.getItem(self, "选择文件", "请选择要清理的文件:", 
                                             [os.path.basename(f) for f in self.loaded_files.keys()], 0, False)
        if ok and file_name:
            file_path = next(path for path in self.loaded_files.keys() if os.path.basename(path) == file_name)
            sheet_name, ok = QInputDialog.getItem(self, "选择工作表", "请选择要清理的工作表:", 
                                                  list(self.loaded_files[file_path].keys()), 0, False)
            if ok and sheet_name:
                df = self.loaded_files[file_path][sheet_name]['data']
                
                # 执行数据清理操作
                df = df.dropna()  # 删除包含空值的行
                df = df.drop_duplicates()  # 删除重复行
                
                # 更新已加载的文件数据
                self.loaded_files[file_path][sheet_name]['data'] = df
                
                self.log(f"已清理文件 {file_name} 的 {sheet_name} 工作表")
                QMessageBox.information(self, "清理完成", f"已成功清理 {file_name} 的 {sheet_name} 工作表")
                
                # 更新预览
                self.preview_file(self.file_list.findItems(file_name, Qt.MatchFlag.MatchExactly)[0])

    def show_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.load_settings()

    def show_help(self):
        help_dialog = HelpDialog(self)
        help_dialog.exec()

    def show_about(self):
        QMessageBox.about(self, "关于", f"高级VLOOKUP工具 {self.current_version}\n\n"
                                       "作者: Kilon\n"
                                       "联系方式: a15607467772@163.com\n\n"
                                       "本工具旨在提供高效的VLOOKUP功能，支持多表查找和数据匹配。")

    def check_for_updates(self):
        self.updater.check_for_updates()

    def update_application(self, version):
        QMessageBox.information(self, "更新", f"正在更新到版本 {version}...")
        self.log(f"已更新到版本 {version}")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                self.load_file(file_path)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '确认退出', 
                                     "是否要退出程序？\n未保存的数据将丢失。",
                                     QMessageBox.StandardButton.Yes | 
                                     QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()

    def detect_header_row(self, df, max_rows=10):
        scores = []
        common_header_words = ['id', 'name', 'date', 'time', 'value', 'code', 'type', 'category', 'description']
        
        for i in range(min(max_rows, len(df))):
            row = df.iloc[i]
            score = 0
        
            # 对第一行和第二行给予额外分数
            if i == 0:
                score += 2
            elif i == 1:
                score += 1.5  # 给第二行稍微低一点的额外分数
        
            # 检查字符串的比例
            string_ratio = row.apply(lambda x: isinstance(x, str)).mean()
            score += string_ratio * 2
        
            # 检查空值的比例
            non_null_ratio = 1 - row.isnull().mean()
            score += non_null_ratio
        
            # 检查数据类型的一致性
            dtype_consistency = len(set(row.apply(type))) / len(row)
            score += (1 - dtype_consistency)
        
            # 检查长度的一致性和偏好短字符串
            lengths = row.apply(lambda x: len(str(x)) if x is not None else 0)
            length_consistency = 1 - (lengths.std() / lengths.mean() if lengths.mean() > 0 else 0)
            score += length_consistency
            score += 1 / (lengths.mean() + 1)  # 偏好短字符串
        
            # 检查特殊字符或乱码
            special_char_ratio = row.apply(lambda x: sum(not c.isalnum() and not c.isspace() for c in str(x)) / len(str(x)) if x is not None else 0).mean()
            score -= special_char_ratio
        
            # 检查是否包含常见的列名关键词
            lower_row = row.astype(str).str.lower()
            keyword_match = any(lower_row.str.contains(word).any() for word in common_header_words)
            score += 2 if keyword_match else 0
        
            # 检查是否为连续的数字行（可能是数据而不是列名）
            if row.dtype.name.startswith('int') or row.dtype.name.startswith('float'):
                score -= 1
        
            # 检查是否所有单元格都不为空
            if not row.isnull().any():
                score += 0.5
        
            # 检查是否有重复值（列名通常不会重复）
            if row.nunique() == len(row):
                score += 0.5
        
            scores.append(score)
    
        best_row = scores.index(max(scores))
        return best_row

    def update_sheet_header(self, sheet_name, file_path, index):
        if index == 0:  # 使用智能检测的结果
            header_row = self.loaded_files[file_path][sheet_name]['detected_header']
        else:
            header_row = index - 1
        
        # 更新DataFrame的表头
        df = self.loaded_files[file_path][sheet_name]['data']
        df.columns = df.iloc[header_row]
        df = df.drop(df.index[header_row]).reset_index(drop=True)
        self.loaded_files[file_path][sheet_name]['data'] = df
        
        # 更新相关的UI元素
        self.update_main_column_combo()
        self.update_lookup_column_combos()
        self.update_return_columns()

    def load_version(self):
        config = configparser.ConfigParser()
        version_file = 'version.ini'
        if os.path.exists(version_file):
            config.read(version_file)
            return config.get('VERSION', 'current', fallback='v0.0.0')
        else:
            return 'v0.0.0'

class VLOOKUPThread(QThread):
    progress_update = pyqtSignal(int)
    result_ready = pyqtSignal(pd.DataFrame)
    error_occurred = pyqtSignal(str)

    def __init__(self, main_df, main_column, lookup_tables, return_columns):
        super().__init__()
        self.main_df = main_df
        self.main_column = main_column
        self.lookup_tables = lookup_tables
        self.return_columns = return_columns

    def run(self):
        try:
            result_df = self.main_df.copy()
            total_steps = len(self.lookup_tables)
            
            for i, (lookup_df, lookup_column) in enumerate(self.lookup_tables):
                result_df[self.main_column] = result_df[self.main_column].astype(str)
                lookup_df[lookup_column] = lookup_df[lookup_column].astype(str)
                
                result_df = pd.merge(result_df, 
                                     lookup_df,
                                     left_on=self.main_column,
                                     right_on=lookup_column,
                                     how='left',
                                     suffixes=('', f'_table{i}'))
                
                result_df = result_df.fillna('N/A')
                
                self.progress_update.emit(int((i + 1) / total_steps * 100))

            columns_to_keep = [self.main_column] + [col for col in result_df.columns if col in self.return_columns]
            result_df = result_df[columns_to_keep]

            self.result_ready.emit(result_df)
        except Exception as e:
            self.error_occurred.emit(str(e))
        finally:
            self.progress_update.emit(100)

class RecentFilesDialog(QDialog):
    def __init__(self, recent_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle("最近使用的文件")
        self.setMinimumSize(500, 400)
        self.recent_files = recent_files
        self.selected_files = []
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # 搜索栏
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索文件...")
        self.search_input.textChanged.connect(self.filter_files)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        # 文件列表
        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        # 底部按钮
        button_layout = QHBoxLayout()
        self.clear_all_button = QPushButton("清除所有历史记录")
        self.clear_all_button.clicked.connect(self.clear_all_history)
        button_layout.addWidget(self.clear_all_button)
        button_layout.addStretch()
        self.open_button = QPushButton("打开选中文件")
        self.open_button.clicked.connect(self.open_selected_files)
        button_layout.addWidget(self.open_button)
        self.close_button = QPushButton("关闭")
        self.close_button.clicked.connect(self.close)
        button_layout.addWidget(self.close_button)
        layout.addLayout(button_layout)

        self.populate_file_list()

    def populate_file_list(self):
        self.file_list.clear()
        for file_path, access_time in self.recent_files:
            item = QListWidgetItem()
            checkbox = QCheckBox(f"{os.path.basename(file_path)} - {access_time.strftime('%Y-%m-%d %H:%M:%S')}")
            item.setData(Qt.ItemDataRole.UserRole, file_path)  # 将文件路径存储在 QListWidgetItem 中
            self.file_list.addItem(item)
            self.file_list.setItemWidget(item, checkbox)

    def filter_files(self, text):
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            checkbox = self.file_list.itemWidget(item)
            item.setHidden(text.lower() not in checkbox.text().lower())

    def open_selected_files(self):
        self.selected_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            checkbox = self.file_list.itemWidget(item)
            if checkbox.isChecked():
                self.selected_files.append(item.data(Qt.ItemDataRole.UserRole))  # 从 QListWidgetItem 获取文件路径
        if self.selected_files:
            self.accept()
        else:
            QMessageBox.warning(self, "警告", "请选择至少一个文件")

    def clear_all_history(self):
        reply = QMessageBox.question(self, '确认', '确定要清除所有历史记录吗？',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.recent_files.clear()
            self.parent().clear_recent_files()
            self.populate_file_list()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 设置应用程序图标
    icon_path = os.path.join(os.path.dirname(__file__), 'VLookUp.ico')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    window = AdvancedVLOOKUPTool()
    window.show()
    sys.exit(app.exec())