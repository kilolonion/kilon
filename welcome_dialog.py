# welcome_dialog.py

from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QCheckBox, QDialogButtonBox

class WelcomeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("欢迎使用高级VLOOKUP工具")
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("欢迎使用高级VLOOKUP工具！\n\n本工具提供以下功能："))
        features = [
            "多表VLOOKUP操作",
            "文件预览和管理",
            "数据清理",
            "结果导出（Excel, CSV，PDF）"
        ]
        for feature in features:
            layout.addWidget(QLabel(f"• {feature}"))
        
        self.dont_show_again = QCheckBox("不再显示此对话框")
        layout.addWidget(self.dont_show_again)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

def show_welcome_dialog(parent, settings):
    if not settings.value("hide_welcome", False, type=bool):
        welcome = WelcomeDialog(parent)
        if welcome.exec() == QDialog.DialogCode.Accepted and welcome.dont_show_again.isChecked():
            settings.setValue("hide_welcome", True)