import os
import requests
import zipfile
import urllib.request
from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtCore import QObject, pyqtSignal

class Updater(QObject):
    update_available = pyqtSignal(str)
    update_progress = pyqtSignal(int)
    update_completed = pyqtSignal()
    update_error = pyqtSignal(str)

    def __init__(self, current_version, update_url):
        super().__init__()
        self.current_version = current_version
        self.update_url = update_url

    def check_for_updates(self):
        try:
            response = requests.get(self.update_url)
            response.raise_for_status()
            latest_version = response.json()['tag_name']
            if latest_version > self.current_version:
                self.update_available.emit(latest_version)
            return latest_version
        except requests.RequestException as e:
            self.update_error.emit(f"检查更新失败: {str(e)}")
            return None

    def download_update(self, version):
        try:
            download_url = f"https://example.com/downloads/AdvancedVLOOKUPTool-{version}.zip"
            download_path = os.path.join(os.path.dirname(__file__), f"update-{version}.zip")
            
            def progress_hook(count, block_size, total_size):
                percent = int(count * block_size * 100 / total_size)
                self.update_progress.emit(percent)

            urllib.request.urlretrieve(download_url, download_path, progress_hook)
            return download_path
        except Exception as e:
            self.update_error.emit(f"下载更新失败: {str(e)}")
            return None

    def install_update(self, download_path):
        try:
            install_dir = os.path.dirname(__file__)
            with zipfile.ZipFile(download_path, 'r') as zip_ref:
                zip_ref.extractall(install_dir)
            os.remove(download_path)
            self.update_completed.emit()
        except Exception as e:
            self.update_error.emit(f"安装更新失败: {str(e)}")

    def update_application(self, version):
        download_path = self.download_update(version)
        if download_path:
            self.install_update(download_path)

def show_update_dialog(parent, version):
    reply = QMessageBox.question(
        parent,
        '更新可用',
        f"发现新版本 {version}，是否更新？",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.Yes
    )
    return reply == QMessageBox.StandardButton.Yes

def show_update_completed_dialog(parent):
    QMessageBox.information(
        parent,
        "更新完成",
        "更新已完成，请重启应用程序以应用更新。"
    )