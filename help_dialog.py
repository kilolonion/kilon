from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTabWidget, QTextEdit, QPushButton, QScrollArea, QWidget
from PyQt6.QtCore import Qt

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("高级VLOOKUP工具帮助")
        self.setGeometry(200, 200, 800, 600)

        layout = QVBoxLayout(self)
        
        tab_widget = QTabWidget()
        
        # 使用说明标签页
        instructions = QTextEdit()
        instructions.setReadOnly(True)
        instructions.setHtml(self.get_instructions_html())
        tab_widget.addTab(instructions, "使用说明")
        
        # FAQ标签页
        faq = QTextEdit()
        faq.setReadOnly(True)
        faq.setHtml(self.get_faq_html())
        tab_widget.addTab(faq, "常见问题")
        
        # 功能介绍标签页
        features = QTextEdit()
        features.setReadOnly(True)
        features.setHtml(self.get_features_html())
        tab_widget.addTab(features, "功能介绍")
        
        layout.addWidget(tab_widget)
        
        close_button = QPushButton("关闭")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button)

    def get_instructions_html(self):
        return """
        <h2>使用说明</h2>
        <h3>1. 文件管理</h3>
        <ul>
            <li><strong>加载文件：</strong> 点击"选择文件"按钮或直接将文件拖放到应用程序窗口中。支持 .xlsx 和 .xls 格式。</li>
            <li><strong>最近使用的文件：</strong> 点击"最近使用的文件"按钮可以快速访问之前打开过的文件。</li>
            <li><strong>预览文件：</strong> 在左侧文件列表中选择文件，可以在下方预览表格内容。</li>
            <li><strong>删除文件：</strong> 选中文件后点击"删除文件"按钮可以从当前会话中移除文件。</li>
            <li><strong>清除所有文件：</strong> 使用"文件"菜单中的"清除所有文件"选项可以移除所有已加载的文件。</li>
        </ul>

        <h3>2. 设置VLOOKUP参数</h3>
        <ul>
            <li><strong>选择主表：</strong> 在右侧面板的"选择主表"下拉菜单中选择要作为主表的文件和工作表。</li>
            <li><strong>选择主列：</strong> 在"选择主列"下拉菜单中选择用于匹配的主列。</li>
            <li><strong>选择查找表：</strong> 在"选择查找表"列表中勾选要用于查找的表格。</li>
            <li><strong>选择查找列：</strong> 为每个选中的查找表选择对应的查找列。</li>
            <li><strong>选择返回列：</strong> 在返回列列表中勾选需要包含在结果中的列。可以使用搜索框筛选列名。</li>
            <li><strong>表头选择：</strong> 对于每个加载的表格，您可以选择使用智能检测的表头或手动指定表头行。</li>
        </ul>

        <h3>3. 执行VLOOKUP</h3>
        <ul>
            <li>设置好所有参数后，点击"执行VLOOKUP"按钮开始操作。</li>
            <li>操作进度会显示在进度条中。</li>
            <li>完成后，结果会显示在右下方的结果表格中。</li>
        </ul>

        <h3>4. 结果处理</h3>
        <ul>
            <li><strong>保存结果：</strong> 点击"保存结果"按钮，选择保存格式（Excel 或 CSV）和位置。</li>
            <li><strong>数据清理：</strong> 使用"工具"菜单中的"数据清理"功能可以对选定的表格进行基本的数据清理。</li>
        </ul>

        <h3>5. 其他功能</h3>
        <ul>
            <li><strong>设置：</strong> 在"设置"菜单中可以配置默认保存格式、自动更新检查等选项。</li>
            <li><strong>帮助：</strong> "帮助"菜单提供此帮助文档和关于信息。</li>
            <li><strong>日志：</strong> 操作日志会显示在主界面的底部，记录重要的操作和事件。</li>
        </ul>
        """

    def get_faq_html(self):
        return """
        <h2>常见问题</h2>
        <h3>Q1: 我可以同时处理多少个文件？</h3>
        <p>A: 理论上没有限制，但实际数量取决于您的计算机性能和可用内存。建议同时处理的文件不要超过10个，以保证良好的性能。</p>

        <h3>Q2: 如何在多个查找表上执行VLOOKUP？</h3>
        <p>A: 在"选择查找表"列表中勾选多个表格，然后为每个表格选择对应的查找列。程序会自动对所有选中的表执行VLOOKUP操作，并将结果合并。</p>

        <h3>Q3: 为什么我的某些列没有出现在返回列列表中？</h3>
        <p>A: 返回列列表只显示您选择的查找表中的列。确保您已经选择了正确的查找表，并且这些表中包含您需要的列。</p>

        <h3>Q4: 如何更改表头？程序如何识别表头？</h3>
        <p>A: 程序会自动尝试检测表头，但您也可以手动选择。在加载文件后，每个表格旁边会有一个下拉菜单，您可以使用它来选择正确的表头行。程序使用启发式算法来检测最可能的表头行，考虑因素包括字符串比例、空值比例、数据类型一致性等。</p>

        <h3>Q5: 数据清理功能具体会做什么？</h3>
        <p>A: 当前的数据清理功能主要执行两项操作：1) 删除包含空值的行；2) 删除重复行。</p>

        <h3>Q6: 如何更新程序？</h3>
        <p>A: 程序默认会在启动时自动检查更新。如果有新版本可用，您会收到通知。您也可以在"设置"中配置更新选项，或禁用自动更新检查。</p>

        <h3>Q7: 为什么我的VLOOKUP结果中有"N/A"值？</h3>
        <p>A: "N/A"值表示在查找表中没有找到匹配的值。这可能是因为主表中的某些值在查找表中不存在，或者数据类型不一致（例如，一个是文本，一个是数字）。</p>

        <h3>Q8: 程序支持哪些文件格式？</h3>
        <p>A: 目前程序支持 .xlsx 和 .xls 格式的 Excel 文件。对于结果保存，还支持 .csv 格式。</p>

        <h3>Q9: 如何报告bug或请求新功能？</h3>
        <p>A: 请发送邮件至 support@advancedvlookuptool.com，或在我们的 GitHub 仓库中创建一个 issue。我们非常重视用户反馈，并会尽快回应。</p>

        <h3>Q10: 为什么我无法选择某些列作为主列或查找列？</h3>
        <p>A: 确保您已经正确选择了表头。如果列名显示为数字，可能是因为程序没有正确识别表头。尝试使用表头选择下拉菜单来手动选择正确的表头行。</p>
        """

    def get_features_html(self):
        return """
        <h2>功能介绍</h2>
        <h3>1. 多文件管理</h3>
        <p>支持同时加载和管理多个Excel文件。您可以通过文件选择对话框或拖放操作来加载文件。程序会记住最近使用的文件，方便快速访问。</p>

        <h3>2. 智能表头检测</h3>
        <p>使用先进的算法自动检测表头行。考虑因素包括字符串比例、空值比例、数据类型一致性、特殊字符存在性、常见列名关键词匹配等。用户也可以手动选择表头行。</p>

        <h3>3. 多表VLOOKUP</h3>
        <p>支持在多个查找表上同时执行VLOOKUP操作。您可以选择多个查找表，程序会自动合并所有查找结果。</p>

        <h3>4. 灵活的列选择</h3>
        <p>为主表和每个查找表分别选择匹配列。可以自由选择要包含在结果中的返回列，支持列名搜索和过滤。</p>

        <h3>5. 数据预览</h3>
        <p>加载文件后可以预览表格内容，帮助您确认数据正确性和选择合适的列。</p>

        <h3>6. 数据清理</h3>
        <p>提供基本的数据清理功能，包括删除空值行和重复行。</p>

        <h3>7. 多格式结果导出</h3>
        <p>支持将结果保存为Excel (.xlsx) 和 CSV 格式。您可以在设置中配置默认的保存格式。</p>

        <h3>8. 最近文件记录</h3>
        <p>自动记录最近使用的文件，并提供快速访问。您可以在设置中配置要记住的文件数量。</p>

        <h3>9. 拖放支持</h3>
        <p>支持通过拖放操作快速加载文件，提高工作效率。</p>

        <h3>10. 自动更新</h3>
        <p>程序会自动检查和提示可用的更新。您可以在设置中配置更新检查的频率或禁用自动更新检查。</p>

        <h3>11. 自定义设置</h3>
        <p>提供多项可自定义的设置，如默认保存格式、更新检查频率、最大记住的文件数等。</p>

        <h3>12. 日志记录</h3>
        <p>详细记录操作过程，方便追踪和调试。日志包含操作时间、类型和结果等信息，显示在主界面底部。</p>

        <h3>13. 用户友好界面</h3>
        <p>直观的图形用户界面，清晰的布局和操作流程，适合各级用户使用。提供详细的帮助文档和常见问题解答。</p>

        <h3>14. 错误处理和反馈</h3>
        <p>强大的错误处理机制，为用户提供清晰的错误信息和可能的解决方案。</p>

        <h3>15. 多线程处理</h3>
        <p>VLOOKUP操作在单独的线程中执行，避免界面卡顿，提供更好的用户体验。</p>

        <h3>未来计划</h3>
        <p>我们计划在未来版本中添加更多功能，如支持更多文件格式、更高级的数据清理选项、图表生成功能、批处理功能等。我们非常重视用户反馈，欢迎提供您的建议和需求。</p>
        """