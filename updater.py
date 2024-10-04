import os
import sys
import requests
import zipfile
import urllib.request
import shutil
import subprocess
from PyQt6.QtWidgets import QMessageBox, QProgressBar
from PyQt6.QtCore import QObject, pyqtSignal
import configparser

class Updater(QObject):
    update_available = pyqtSignal(str)
    update_progress = pyqtSignal(int)  # 将这个改回信号
    update_completed = pyqtSignal()
    update_error = pyqtSignal(str)

    def __init__(self, current_version, update_url):
        super().__init__()
        self.current_version = current_version
        self.update_url = update_url
        self.latest_version = None
        self.progress_bar = None

    def check_for_updates(self):
        try:
            response = requests.get(self.update_url)
            response.raise_for_status()
            self.latest_version = response.json()['tag_name']
            if self.latest_version > self.current_version:
                self.update_available.emit(self.latest_version)
            return self.latest_version
        except requests.RequestException as e:
            self.update_error.emit(f"检查更新失败: {str(e)}")
            return None

    def download_update(self, version):
        try:
            download_url = f"https://github.com/kilolonion/kilon/archive/refs/tags/{version}.zip"
            download_path = os.path.join(os.path.dirname(__file__), f"update-{version}.zip")
            
            def progress_hook(count, block_size, total_size):
                percent = int(count * block_size * 100 / total_size)
                self.update_progress_bar(percent)  # 使用新的方法名

            urllib.request.urlretrieve(download_url, download_path, progress_hook)
            return download_path
        except Exception as e:
            self.update_error.emit(f"下载更新失败: {str(e)}")
            return None

    def install_update(self, download_path):
        try:
            install_dir = os.path.dirname(__file__)
            temp_dir = os.path.join(install_dir, 'temp_update')
            
            # 解压到临时目录
            with zipfile.ZipFile(download_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 获取解压后的文件夹名称
            extracted_folder = os.path.join(temp_dir, os.listdir(temp_dir)[0])
            
            # 替换旧文件
            for item in os.listdir(extracted_folder):
                s = os.path.join(extracted_folder, item)
                d = os.path.join(install_dir, item)
                if os.path.isdir(s):
                    shutil.copytree(s, d, dirs_exist_ok=True)
                else:
                    shutil.copy2(s, d)
            
            # 清理临时文件
            os.remove(download_path)
            shutil.rmtree(temp_dir)
            
            # 更新 version.ini 文件
            self.update_version_file(install_dir)
            
            self.update_completed.emit()
            
            # 重启应用
            self.restart_application()
            
        except Exception as e:
            self.update_error.emit(f"安装更新失败: {str(e)}")

    def update_version_file(self, install_dir):
        config = configparser.ConfigParser()
        version_file = os.path.join(install_dir, 'version.ini')
        config.read(version_file)
        if self.latest_version:
            config.set('VERSION', 'current', self.latest_version)
            with open(version_file, 'w') as configfile:
                config.write(configfile)
        else:
            raise ValueError("最新版本号未设置")

    def restart_application(self):
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def update_application(self, version):
        self.latest_version = version
        download_path = self.download_update(version)
        if download_path:
            self.install_update(download_path)

    def show_update_dialog(self, parent, version):
        reply = QMessageBox.question(
            parent,
            '更新可用',
            f"发现新版本 {version}，是否更新？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )
        return reply == QMessageBox.StandardButton.Yes

    def show_update_completed_dialog(self, parent):
        QMessageBox.information(
            parent,
            "更新完成",
            "更新已完成，应用程序将重新启动以应用更新。"
        )

    def set_progress_bar(self, progress_bar):
        self.progress_bar = progress_bar

    def update_progress_bar(self, percent):  # 重命名这个方法
        if self.progress_bar:
            self.progress_bar.setValue(percent)
        self.update_progress.emit(percent)

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