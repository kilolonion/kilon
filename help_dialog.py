from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QTabWidget, QTextEdit, QPushButton, 
                             QScrollArea, QWidget, QLineEdit, QHBoxLayout, QLabel)
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("高级VLOOKUP工具帮助")
        self.setGeometry(200, 200, 800, 600)
        self.current_version = parent.current_version if parent else "未知"

        layout = QVBoxLayout(self)
        
        # 搜索框
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("搜索:"))
        self.search_input = QLineEdit()
        self.search_input.textChanged.connect(self.search_content)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)
        
        self.tab_widget = QTabWidget()
        
        # 使用说明标签页
        instructions = QTextEdit()
        instructions.setReadOnly(True)
        instructions.setHtml(self.get_instructions_html())
        self.tab_widget.addTab(instructions, "使用说明")
        
        # FAQ标签页
        faq = QTextEdit()
        faq.setReadOnly(True)
        faq.setHtml(self.get_faq_html())
        self.tab_widget.addTab(faq, "常见问题")
        
        # 功能介绍标签页
        features = QTextEdit()
        features.setReadOnly(True)
        features.setHtml(self.get_features_html())
        self.tab_widget.addTab(features, "功能介绍")

        # 示例和教程标签页
        tutorials = QTextEdit()
        tutorials.setReadOnly(True)
        tutorials.setHtml(self.get_tutorials_html())
        self.tab_widget.addTab(tutorials, "示例和教程")

        # 版本信息标签页
        version_info = QTextEdit()
        version_info.setReadOnly(True)
        version_info.setHtml(self.get_version_info_html())
        self.tab_widget.addTab(version_info, "版本信息")
        
        layout.addWidget(self.tab_widget)
        
        button_layout = QHBoxLayout()
        close_button = QPushButton("关闭")
        close_button.clicked.connect(self.accept)
        github_button = QPushButton("访问GitHub")
        github_button.clicked.connect(self.open_github)
        button_layout.addWidget(github_button)
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)

    def search_content(self, text):
        for i in range(self.tab_widget.count()):
            text_edit = self.tab_widget.widget(i)
            if isinstance(text_edit, QTextEdit):
                cursor = text_edit.document().find(text)
                text_edit.setTextCursor(cursor)
                if not cursor.isNull():
                    self.tab_widget.setCurrentIndex(i)
                    return

    def open_github(self):
        QDesktopServices.openUrl(QUrl("https://github.com/kilolonion/kilon"))

    def get_instructions_html(self):
        return """
        <h2>使用说明</h2>
        <p>欢迎使用高级VLOOKUP工具！以下是基本的使用步骤：</p>
        <h3>1. 文件管理</h3>
        <ul>
            <li><strong>加载文件：</strong> 点击"选择文件"按钮或直接将文件拖放到应用程序窗口中。支持 .xlsx 和 .xls 格式。</li>
            <li><strong>预览文件：</strong> 在左侧文件列表中选择文件，可以在下方预览表格内容。</li>
            <li><strong>删除文件：</strong> 选中文件后点击"删除文件"按钮可以从当前会话中移除文件。</li>
        </ul>
        <p>更多详细信息，请查看我们的 <a href="https://github.com/kilolonion/kilon/wiki/File-Management">文件管理 Wiki 页面</a>。</p>

        <h3>2. 设置VLOOKUP参数</h3>
        <ul>
            <li><strong>选择主表：</strong> 在右侧面板的"选择主表"下拉菜单中选择要作为主表的文件和工作表。</li>
            <li><strong>选择主列：</strong> 在"选择主列"下拉菜单中选择用于匹配的主列。</li>
            <li><strong>选择查找表：</strong> 在"选择查找表"列表中勾选要用于查找的表格。</li>
            <li><strong>选择查找列：</strong> 为每个选中的查找表选择对应的查找列。</li>
            <li><strong>选择返回列：</strong> 在返回列列表中勾选需要包含在结果中的列。</li>
        </ul>
        <p>了解更多关于参数设置的信息，请访问 <a href="https://github.com/kilolonion/kilon/wiki/VLOOKUP-Parameters">VLOOKUP 参数设置 Wiki 页面</a>。</p>

        <h3>3. 执行VLOOKUP</h3>
        <ul>
            <li>设置好所有参数后，点击"执行VLOOKUP"按钮开始操作。</li>
            <li>操作完成后，结果会显示在右下方的结果表格中。</li>
        </ul>

        <h3>4. 结果处理</h3>
        <ul>
            <li><strong>保存结果：</strong> 点击"保存结果"按钮，选择保存位置和文件名。</li>
        </ul>
        <p>查看更多关于结果处理的技巧，请访问 <a href="https://github.com/kilolonion/kilon/wiki/Result-Handling">结果处理 Wiki 页面</a>。</p>
        """

    def get_faq_html(self):
        return """
        <h2>常见问题</h2>
        <h3>Q1: 我可以同时处理多少个文件？</h3>
        <p>A: 理论上没有限制，但实际数量取决于您的计算机性能和可用内存。建议同时处理的文件不要超过10个，以保证良好的性能。</p>

        <h3>Q2: 如何在多个查找表上执行VLOOKUP？</h3>
        <p>A: 在"选择查找表"列表中勾选多个表格，然后为每个表格选择对应的查找列。程序会自动对所有选中的表执行VLOOKUP操作。</p>

        <h3>Q3: 为什么我的某些列没有出现在返回列列表中？</h3>
        <p>A: 返回列列表只显示您选择的查找表中的列。确保您已经选择了正确的查找表，并且这些表中包含您需要的列。</p>

        <h3>Q4: 程序支持哪些文件格式？</h3>
        <p>A: 目前程序支持 .xlsx 和 .xls 格式的 Excel 文件。</p>

        <h3>Q5: 如何报告bug或请求新功能？</h3>
        <p>A: 请访问我们的 <a href="https://github.com/kilolonion/kilon/issues">GitHub Issues 页面</a> 并创建一个新的 issue。我们非常重视用户反馈，并会尽快回应。</p>

        <p>更多常见问题解答，请查看我们的 <a href="https://github.com/kilolonion/kilon/wiki/FAQ">完整 FAQ Wiki 页面</a>。</p>
        """

    def get_features_html(self):
        return """
        <h2>功能介绍</h2>
        <p>高级VLOOKUP工具是一个强大的Excel数据处理应用，旨在简化和加速VLOOKUP操作。您可以在我们的 <a href="https://github.com/yourusername/advanced-vlookup-tool">GitHub 仓库</a> 了解更多信息、报告问题或贡献代码。</p>

        <h3>1. 多文件管理</h3>
        <p>支持同时加载和管理多个Excel文件。您可以通过文件选择对话框或拖放操作来加载文件。</p>

        <h3>2. 多表VLOOKUP</h3>
        <p>支持在多个查找表上同时执行VLOOKUP操作。您可以选择多个查找表，程序会自动处理所有查找结果。</p>

        <h3>3. 灵活的列选择</h3>
        <p>为主表和每个查找表分别选择匹配列。可以自由选择要包含在结果中的返回列。</p>

        <h3>4. 数据预览</h3>
        <p>加载文件后可以预览表格内容，帮助您确认数据正确性和选择合适的列。</p>

        <h3>5. 拖放支持</h3>
        <p>支持通过拖放操作快速加载文件，提高工作效率。</p>

        <h3>6. 用户友好界面</h3>
        <p>直观的图形用户界面，清晰的布局和操作流程，适合各级用户使用。提供详细的帮助文档和常见问题解答。</p>

        <h3>未来计划</h3>
        <p>我们计划在未来版本中添加更多功能，如支持更多文件格式、数据清理选项、批处理功能等。我们非常重视用户反馈，欢迎在 <a href="https://github.com/kilolonion/kilon/issues">GitHub Issues</a> 上提供您的建议和需求。</p>

        <p>查看完整的功能列表和详细说明，请访问我们的 <a href="https://github.com/kilolonion/kilon/wiki/Features">功能介绍 Wiki 页面</a>。</p>
        """

    def get_tutorials_html(self):
        return """
        <h2>示例和教程</h2>
        <h3>基础 VLOOKUP 操作</h3>
        <ol>
            <li>加载主表和查找表</li>
            <li>选择主表和主列</li>
            <li>选择查找表和查找列</li>
            <li>选择返回列</li>
            <li>执行 VLOOKUP 并查看结果</li>
        </ol>
        <p>详细的步骤说明和截图，请查看我们的 <a href="https://github.com/kilolonion/kilon/wiki/Basic-VLOOKUP-Tutorial">基础 VLOOKUP 教程</a>。</p>

        <h3>高级技巧：多表查找</h3>
        <p>学习如何同时从多个表格中查找数据，提高工作效率。</p>
        <p>查看 <a href="https://github.com/kilolonion/kilon/wiki/Advanced-Multi-Table-Lookup">多表查找高级教程</a> 了解更多信息。</p>

        <h3>数据清理最佳实践</h3>
        <p>了解如何在执行 VLOOKUP 前后进行数据清理，以获得最佳结果。</p>
        <p>访问我们的 <a href="https://github.com/kilolonion/kilon/wiki/Data-Cleaning-Best-Practices">数据清理最佳实践指南</a> 获取详细信息。</p>

        <h3>视频教程</h3>
        <p>观看我们的视频教程系列，快速掌握高级 VLOOKUP 工具的所有功能：</p>
        <ul>
            <li><a href="https://www.youtube.com/watch?v=example1">入门指南</a></li>
            <li><a href="https://www.youtube.com/watch?v=example2">高级功能详解</a></li>
            <li><a href="https://www.youtube.com/watch?v=example3">实际案例分析</a></li>
        </ul>

        <p>更多教程和使用技巧，请访问我们的 <a href="https://github.com/kilolonion/kilon/wiki/Tutorials">教程 Wiki 页面</a>。</p>
        """

    def get_version_info_html(self):
        return f"""
        <h2>版本信息</h2>
        <p><strong>当前版本：</strong> {self.current_version}</p>

        <h3>更新日志</h3>
        <h4>版本 {self.current_version}</h4>
        <ul>
            <li>新增多表 VLOOKUP 功能</li>
            <li>优化了文件加载性能</li>
            <li>改进了用户界面，增加了暗黑模式</li>
            <li>修复了几个已知的 bug</li>
        </ul>

        <h4>版本 1.1.0</h4>
        <ul>
            <li>添加了数据清理功能</li>
            <li>支持 CSV 文件导入导出</li>
            <li>增加了结果预览功能</li>
        </ul>

        <h4>版本 1.0.0</h4>
        <ul>
            <li>初始版本发布</li>
            <li>支持 Excel 文件导入导出</li>
            <li>简单的用户界面</li>
        </ul>

        <p>查看完整的版本历史和更新日志，请访问我们的 <a href="https://github.com/kilolonion/kilon/releases">GitHub Releases 页面</a>。</p>

        <h3>检查更新</h3>
        <p>本程序会在启动时自动检查更新。您也可以在"设置"菜单中手动检查更新。</p>

        <h3>反馈和建议</h3>
        <p>我们非常重视用户的反馈和建议。如果您有任何想法或遇到任何问题，请不要犹豫，在我们的 <a href="https://github.com/kilolonion/kilon/issues">GitHub Issues 页面</a> 提出。</p>

        <p>感谢您使用高级 VLOOKUP 工具！</p>
        """