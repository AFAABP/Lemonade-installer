import os
import requests
from PyQt6.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QCheckBox, QStackedLayout, QHBoxLayout, QGroupBox, QComboBox, QProgressBar, QMessageBox
from PyQt6.QtGui import QPixmap, QIcon, QImage
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QByteArray
import zipfile
import shutil
import tempfile
import win32com.client
import logging
import re
import winreg as reg
import base64
from image_base64 import image_data


class DownloadThread(QThread):
    progress = pyqtSignal(int)

    def __init__(self, url, dest):
        super().__init__()
        self.url = url
        self.dest = dest

    def run(self):
        try:
            response = requests.get(self.url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            if total_size == 0:
                QMessageBox.critical(self, "Error", "The content-length of the response is zero.")
                return

            downloaded_size = 0
            with open(self.dest, 'wb') as file:
                for data in response.iter_content(1024):
                    downloaded_size += len(data)
                    file.write(data)
                    progress_percentage = (downloaded_size / total_size) * 100
                    self.progress.emit(int(progress_percentage))
            print(f"Download completed. File saved to {self.dest}")
        except Exception as e:
            QMessageBox.critical(self, "Error", "Error doing downlaod.")

class Installer(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QStackedLayout()
        self.setWindowIcon(QIcon('lemonade.ico'))
        self.setWindowTitle('Lemonade Installer')  # Set the window title

        # Function to create a header with the icon
        def createHeader():
            headerLayout = QHBoxLayout()
            headerLayout.addStretch(1)
            iconLabel = QLabel()
            image_bytes = base64.b64decode(image_data)
            image = QImage.fromData(QByteArray(image_bytes))
            pixmap = QPixmap.fromImage(image)
            # Scale the pixmap to a new size while maintaining the aspect ratio
            scaledPixmap = pixmap.scaled(128, 128, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            iconLabel.setPixmap(scaledPixmap)
            headerLayout.addWidget(iconLabel)
            headerLayout.addStretch(1)
            return headerLayout

        # Disclaimer page with QGroupBox
        self.disclaimerPage = QWidget()
        disclaimerLayout = QVBoxLayout()
        disclaimerGroup = QGroupBox("")
        disclaimerGroupLayout = QVBoxLayout()
        disclaimerLabel = QLabel('[<b>Lemonade is still on the early alpha stage, Dont expect big noticible chnages in performance yet/]')
        disclaimerGroupLayout.addWidget(disclaimerLabel)
        disclaimerGroupLayout.addWidget(disclaimerLabel)
        disclaimerGroup.setLayout(disclaimerGroupLayout)
        disclaimerLayout.addLayout(createHeader())  # Add the icon to the header
        disclaimerLayout.addWidget(disclaimerGroup)
        nextButton = QPushButton('Next')
        nextButton.clicked.connect(self.showInstallPage)
        disclaimerLayout.addWidget(nextButton)
        self.disclaimerPage.setLayout(disclaimerLayout)

        # Install page with QGroupBox for checkboxes
        self.installPage = QWidget()
        installLayout = QVBoxLayout()
        installLayout.addLayout(createHeader())
        self.label = QLabel('<b>Please select a version and click "Install"')
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        installLayout.addWidget(self.label)

        # Dropdown menu for selecting installation source
        self.installationSourceComboBox = QComboBox()
        self.installationSourceComboBox.addItem("Latest Release")  # Added option for Latest nightly
        self.installationSourceComboBox.addItem("Latest Nightly")  # Added option for Latest Release
        installLayout.addWidget(self.installationSourceComboBox)

        # GroupBox for checkboxes
        checkboxGroup = QGroupBox("Do you want to create shortcuts?")
        checkboxLayout = QVBoxLayout()
        self.desktopShortcutCheckbox = QCheckBox("Create a desktop shortcut")
        self.startMenuShortcutCheckbox = QCheckBox("Create a start menu shortcut")
        checkboxLayout.addWidget(self.desktopShortcutCheckbox)
        checkboxLayout.addWidget(self.startMenuShortcutCheckbox)
        checkboxGroup.setLayout(checkboxLayout)

        self.button = QPushButton('Install')
        self.button.clicked.connect(self.prepareAndInstall)

        # Add widgets to the layout
        installLayout.addWidget(checkboxGroup)
        installLayout.addWidget(self.button)
        self.installPage.setLayout(installLayout)

        # Add pages to the stack and set layout
        self.layout.addWidget(self.disclaimerPage)
        self.layout.addWidget(self.installPage)
        self.setLayout(self.layout)

        # Page for the progress bars
        self.progressBarPage = QWidget()
        progressBarLayout = QVBoxLayout()

        # Add the icon to the progress bar page
        iconLabel = QLabel()
        image_bytes = base64.b64decode(image_data)
        image = QImage.fromData(QByteArray(image_bytes))
        pixmap = QPixmap.fromImage(image)
        scaledPixmap = pixmap.scaled(128, 128, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        iconLabel.setPixmap(scaledPixmap)
        progressBarLayout.addWidget(iconLabel, alignment=Qt.AlignmentFlag.AlignCenter)  # Add the icon and center it

        self.downloadProgressBar = QProgressBar()
        progressBarLayout.addStretch(1)
        self.downloadProgressBar.setRange(0, 100)  # Assuming 0 to 100% progress
        progressBarLayout.addWidget(QLabel("Downloading Lemonade:"))
        progressBarLayout.addWidget(self.downloadProgressBar)

        self.extractionProgressBar = QProgressBar()
        self.extractionProgressBar.setRange(0, 100)  # Assuming 0 to 100% progress
        progressBarLayout.addWidget(QLabel("Extracting Lemonade:"))
        progressBarLayout.addWidget(self.extractionProgressBar)

        self.progressBarPage.setLayout(progressBarLayout)
        self.layout.addWidget(self.progressBarPage)

        # Add the icon to the finish page
        iconLabel = QLabel()
        image_bytes = base64.b64decode(image_data)
        image = QImage.fromData(QByteArray(image_bytes))
        pixmap = QPixmap.fromImage(image)
        scaledPixmap = pixmap.scaled(128, 128, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        iconLabel.setPixmap(scaledPixmap)

        self.finishPage = QWidget()
        finishLayout = QVBoxLayout()
        finishLabel = QLabel("<b>Installation Complete!")
        finishLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Center the text horizontally
        finishButton = QPushButton("Finish")
        finishButton.clicked.connect(self.close)
        finishLayout.addWidget(iconLabel, alignment=Qt.AlignmentFlag.AlignCenter)  # Add the icon and center it
        finishLayout.addWidget(finishLabel)
        finishLayout.addWidget(finishButton)
        self.finishPage.setLayout(finishLayout)

        self.layout.addWidget(self.finishPage)


    def showInstallPage(self):
        self.layout.setCurrentIndex(1)

    def prepareAndInstall(self):
        # Show the progress bar page
        self.showProgressBarPage()
        # Then start the installation process
        self.install()

    def showProgressBarPage(self):
        self.layout.setCurrentIndex(2)


    def install(self):
        selection = self.installationSourceComboBox.currentText()
        if selection == "Latest Nightly":
            url = "https://nightly.link/Lemonade-emu/Lemonade/workflows/build/master/windows-msvc.zip"
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                print("Download successful, proceeding with installation.")
            else:
                QMessageBox.critical(self, "Error", "Failed to download. HTTP status code: {response.status_code}.")
                return
        elif selection == "Latest Release":
            url = self.get_latest_release_url()
            if not url:
                QMessageBox.critical(self, "Error", "Repo ULR was not found.")
                return
        temp_file = tempfile.NamedTemporaryFile(delete=False).name
        self.download_thread = DownloadThread(url, temp_file)
        self.download_thread.progress.connect(self.downloadProgressBar.setValue)
        self.download_thread.finished.connect(lambda: self.extract_zip(temp_file, os.path.join(os.environ['LOCALAPPDATA'], 'Lemonade')))
        self.download_thread.start()

    def get_latest_release_url(self):
        try:
            api_url = "https://api.github.com/repos/Lemonade-emu/Lemonade/releases"
            response = requests.get(api_url)
            if response.status_code != 200:
                QMessageBox.critical(self, "Error", "Failed to fetch releases from GitHub.")
                return None

            releases = response.json()
            for release in releases:
                assets = release.get('assets', [])
                for asset in assets:
                    if "lemonade-windows-msvc" in asset['name']:
                        return asset['browser_download_url']
            QMessageBox.critical(self, "Error", "No suitable release found.")
            return None
        except Exception as e:
            QMessageBox.critical(self, "Exception", f"An error occurred: {e}")
            return None

    def clear_directory(self, directory):
        """
        Removes all files and directories in the specified directory.

        :param directory: Path to the directory to clear.
        """
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            try:
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
            except Exception as e:
                logging.error(f'Failed to delete {item_path}. Reason: {e}')

    def extract_zip(self, temp_file, extract_to):
        try:
            # Clear the target directory before extracting new files
            if os.path.exists(extract_to):
                self.clear_directory(extract_to)

            # Rename the temporary file to have a .zip extension
            zip_file_path = temp_file + '.zip'
            os.rename(temp_file, zip_file_path)

            # Temporary extraction folder
            temp_extract_folder = tempfile.mkdtemp()

            # Extract the main zip file to a temporary folder
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_folder)
                logging.info(f"Main zip extraction completed to temporary folder.")

            # Find the nested zip file or the main directory
            nested_zip_path = None
            for root, dirs, files in os.walk(temp_extract_folder):
                for file in files:
                    if file.endswith('.zip'):
                        nested_zip_path = os.path.join(root, file)
                        break
                if nested_zip_path:  # Stop searching if the nested zip is found
                    break

            if nested_zip_path:
                with zipfile.ZipFile(nested_zip_path, 'r') as nested_zip_ref:
                    nested_zip_ref.extractall(extract_to)
                    logging.info(f"Nested zip extraction completed to {extract_to}.")
            else:
                # If no nested zip, move the contents from the temp folder to the desired location
                for item in os.listdir(temp_extract_folder):
                    s = os.path.join(temp_extract_folder, item)
                    d = os.path.join(extract_to, item)
                    if os.path.isdir(s):
                        shutil.move(s, d)
                    else:
                        shutil.copy2(s, d)
                logging.info("Moved extracted files to the desired location.")

            # After extracting all files
            nested_dir_name = None
            for item in os.listdir(extract_to):
                if os.path.isdir(os.path.join(extract_to, item)) and 'lemonade-windows-msvc' in item:
                    nested_dir_name = item
                    break

            if nested_dir_name:
                nested_dir_path = os.path.join(extract_to, nested_dir_name)
                for item in os.listdir(nested_dir_path):
                    s = os.path.join(nested_dir_path, item)
                    d = os.path.join(extract_to, item)
                    if os.path.isdir(s):
                        shutil.move(s, d)
                    else:
                        shutil.copy2(s, d)
                # Remove the now-empty nested directory
                shutil.rmtree(nested_dir_path)

            # Github info
            release_api_url = "https://api.github.com/repos/Lemonade-emu/Lemonade-installer/releases"
            dest_path = os.path.join(os.environ['LOCALAPPDATA'], 'Lemonade', 'uninstaller.exe')

            # Set the progress bar to 100% and call installation_complete
            self.extractionProgressBar.setValue(100)
            executable_path = os.path.join(os.environ['LOCALAPPDATA'], 'Lemonade', 'lemonade-qt.exe')
            self.download_uninstaller(release_api_url, dest_path)
            self.add_to_programs_list(executable_path)
            self.installation_complete()
        except Exception as e:
            logging.error(f"Failed to extract zip file: {e}. File: {item}")

    def installation_complete(self):
        executable_path = os.path.join(os.environ['LOCALAPPDATA'], 'Lemonade', 'lemonade-qt.exe')
        if self.desktopShortcutCheckbox.isChecked():
            self.create_desktop_shortcut(executable_path)
        if self.startMenuShortcutCheckbox.isChecked():
            self.create_start_menu_shortcut(executable_path)
        self.layout.setCurrentIndex(self.layout.indexOf(self.finishPage))  # Switch to finish page

    def download_uninstaller(self, release_api_url, dest_path):
        try:
            logging.info("Starting to download uninstaller.")
            response = requests.get(release_api_url)
            response.raise_for_status()  # This will raise an exception for HTTP errors
            releases = response.json()
            found = False
            for release in releases:
                for asset in release.get('assets', []):
                    if asset['name'] == "uninstaller.exe":
                        uninstaller_url = asset['browser_download_url']
                        self.download_file(uninstaller_url, dest_path)
                        found = True
                        logging.info("Uninstaller downloaded successfully.")
                        break
                if found:
                    break
            if not found:
                logging.error("Uninstaller not found in any release.")
        except requests.RequestException as e:
            logging.error(f"Error downloading uninstaller: {e}")

    def download_file(self, url, dest_path):
        try:
            with requests.get(url, stream=True) as r:
                r.raise_for_status()
                with open(dest_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
        except requests.RequestException as e:
            print(f"Error saving the uninstaller: {e}")

    def add_to_programs_list(self, executable_path):
        """
        Add the application to the Windows Program list with the uninstall option.

        :param executable_path: Path to the executable file of the application.
        """
        print("Adding to programs list...")
        # Path to the uninstall key
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Lemonade"
        # Uninstaller path
        uninstaller_path = os.path.join(os.environ['LOCALAPPDATA'], 'Lemonade', 'uninstaller.exe')
        print(f"Uninstaller path: {uninstaller_path}")

        # Attempt to open the key, create if it does not exist
        try:
            key = reg.OpenKey(reg.HKEY_CURRENT_USER, key_path, 0, reg.KEY_WRITE)
            print("Registry key exists, opened successfully.")
        except FileNotFoundError:
            key = reg.CreateKey(reg.HKEY_CURRENT_USER, key_path)
            print("Registry key does not exist, created successfully.")

        # Set values within the key
        with key:
            reg.SetValueEx(key, "DisplayName", 0, reg.REG_SZ, "Lemonade")
            reg.SetValueEx(key, "UninstallString", 0, reg.REG_SZ, uninstaller_path)
            reg.SetValueEx(key, "DisplayIcon", 0, reg.REG_SZ, executable_path)
            reg.SetValueEx(key, "Publisher", 0, reg.REG_SZ, "Lemonade-Emu")
            reg.SetValueEx(key, "URLInfoAbout", 0, reg.REG_SZ, "https://lemonade-emu.github.io/")
            print("Registry values set successfully.")

        print("Added to programs list successfully.")

    def create_desktop_shortcut(self, target):
        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        shortcut_path = os.path.join(desktop_path, 'Lemonade.lnk')
        self.create_shortcut(target, shortcut_path)

    def create_start_menu_shortcut(self, target):
        start_menu_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')
        shortcut_path = os.path.join(start_menu_path, 'Lemonade.lnk')
        self.create_shortcut(target, shortcut_path)


    def create_shortcut(self, target, shortcut_path, description="", arguments="", hotkey=""):
        """
        Creates a shortcut at the specified path pointing to the target file.

        :param target: Path to the target file the shortcut will point to.
        :param shortcut_path: Path where the shortcut will be created.
        :param description: Description of the shortcut.
        :param arguments: Additional arguments to pass to the target when executed.
        :param hotkey: Hotkey associated with the shortcut.
        """
        # Verify the target exists
        if not os.path.exists(target):
            logging.error(f"Shortcut target does not exist: {target}")
            return

        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.Description = description
            shortcut.Arguments = arguments
            if hotkey:
                shortcut.Hotkey = hotkey
            shortcut.IconLocation = target  # You can customize this if needed
            shortcut.save()
            logging.info(f"Shortcut created successfully at {shortcut_path}")
        except Exception as e:
            logging.error(f"Failed to create shortcut: {e}")

if __name__ == '__main__':
    app = QApplication([])

    # Define the dark stylesheet
    dark_stylesheet = """
    QWidget {
        background-color: #2b2b2b;
        color: #d8d8d8;
    }
    QPushButton {
        background-color: #3c3f41;
        color: #d8d8d8;
        border-radius: 5px;
        padding: 5px;
    }
    QPushButton:hover {
        background-color: #4b4e51;
    }
    QLabel, QCheckBox {
        color: #d8d8d8;
    }
    QGroupBox {
        border: 1px solid #4b4e51;
        border-radius: 5px;
        margin-top: 20px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 3px 0 3px;
    }
    """

    # Apply the dark stylesheet to the application
    app.setStyleSheet(dark_stylesheet)

    installer = Installer()
    installer.show()
    app.exec()