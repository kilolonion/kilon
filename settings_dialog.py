from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QCheckBox, QPushButton, QLineEdit, QDialogButtonBox)
from PyQt6.QtCore import QSettings

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setGeometry(300, 300, 400, 300)

        self.settings = QSettings("YourCompany", "AdvancedVLOOKUPTool")
        self.current_version = parent.current_version if parent else "未知"
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # 显示当前版本
        version_layout = QHBoxLayout()
        version_layout.addWidget(QLabel("当前版本:"))
        version_layout.addWidget(QLabel(self.current_version))
        layout.addLayout(version_layout)

        log_layout = QHBoxLayout()
        log_layout.addWidget(QLabel("日志级别:"))
        self.log_level_combo = QComboBox()
        self.log_level_combo.addItems(["调试", "信息", "警告", "错误", "严重"])
        log_level_map = {"DEBUG": "调试", "INFO": "信息", "WARNING": "警告", "ERROR": "错误", "CRITICAL": "严重"}
        current_log_level = self.settings.value("log_level", "INFO")
        self.log_level_combo.setCurrentText(log_level_map.get(current_log_level, "信息"))
        log_layout.addWidget(self.log_level_combo)
        layout.addLayout(log_layout)

        save_layout = QHBoxLayout()
        save_layout.addWidget(QLabel("默认保存格式:"))
        self.save_format_combo = QComboBox()
        self.save_format_combo.addItems(["Excel工作簿 (xlsx)", "CSV文件 (csv)", "PDF文档 (pdf)"])
        format_map = {"xlsx": "Excel工作簿 (xlsx)", "csv": "CSV文件 (csv)", "pdf": "PDF文档 (pdf)"}
        current_format = self.settings.value("default_save_format", "xlsx")
        self.save_format_combo.setCurrentText(format_map.get(current_format, "Excel工作簿 (xlsx)"))
        save_layout.addWidget(self.save_format_combo)
        layout.addLayout(save_layout)

        self.auto_update_check = QCheckBox("启动时检查更新")
        self.auto_update_check.setChecked(self.settings.value("auto_update_check", True, type=bool))
        layout.addWidget(self.auto_update_check)

        update_source_layout = QHBoxLayout()
        update_source_layout.addWidget(QLabel("更新源:"))
        self.update_source_input = QLineEdit(self.settings.value("update_source", ""))
        update_source_layout.addWidget(self.update_source_input)
        layout.addLayout(update_source_layout)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("确定")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        layout.addWidget(buttons)

    def accept(self):
        log_level_map = {"调试": "DEBUG", "信息": "INFO", "警告": "WARNING", "错误": "ERROR", "严重": "CRITICAL"}
        self.settings.setValue("log_level", log_level_map[self.log_level_combo.currentText()])
        
        format_map = {"Excel工作簿 (xlsx)": "xlsx", "CSV文件 (csv)": "csv", "PDF文档 (pdf)": "pdf"}
        self.settings.setValue("default_save_format", format_map[self.save_format_combo.currentText()])
        
        self.settings.setValue("auto_update_check", self.auto_update_check.isChecked())
        self.settings.setValue("update_source", self.update_source_input.text())
        super().accept()