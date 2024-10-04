import sys
from PyQt6.QtWidgets import QApplication
from advanced_vlookup_tool import AdvancedVLOOKUPTool

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AdvancedVLOOKUPTool()
    window.show()
    sys.exit(app.exec())