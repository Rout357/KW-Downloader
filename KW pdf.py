from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, NoSuchWindowException, TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service

import sys
import warnings
warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2

from PyQt5.QtWidgets import QApplication, QMainWindow, QSystemTrayIcon, QMenu, QAction, QDialog, QWidget
from PyQt5.QtCore import pyqtSignal
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QRegExpValidator
from PyQt5.QtCore import QRegExp
from PyQt5.QtGui import QIcon

from requests import RequestException
from threading import Thread, Lock
from weasyprint import HTML, CSS
from datetime import datetime

import functools
import pyperclip
import keyboard
import requests

import time
import uuid
import json
import re
import os

import ctypes
import winsound
import tkinter as tk
from tkinter import ttk
import webbrowser


MB_ICONHAND = 0x00000010
MB_OK = 0x00000000



class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setIcon(QIcon('1.png'))
        self.setToolTip('KW pdf')
        self.settings_window_shown = None
        
        menu = QMenu()
        show_action = QAction('Ustawienia', self)
        exit_action = QAction('Zakończ', self)
        menu.addAction(show_action)
        menu.addAction(exit_action)
        self.setContextMenu(menu)
        
        show_action.triggered.connect(self.openSettings)
        exit_action.triggered.connect(QApplication.instance().quit)


    def openSettings(self):
        if self.settings_window_shown is None:
            self.ui = Ui_Settings()
            self.ui.setupUi(self.ui)
            self.ui.setModal(True)
            self.settings_window_shown = self.ui
            self.settings_window_shown.show()
        else:
            self.settings_window_shown.show()


class Ui_Settings(QDialog):
    def __init__(self):
        super().__init__()
        self.settings_file = 'settings.json'
        self.load_settings()
        self.check_path()


    def setupUi(self, Settings):
        Settings.setObjectName("Settings")
        Settings.setWindowModality(QtCore.Qt.WindowModal)
        Settings.setWindowFlags(Settings.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        Settings.resize(450, 430)
        Settings.setMinimumSize(QtCore.QSize(450, 430))
        Settings.setMaximumSize(QtCore.QSize(450, 430))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("1.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Settings.setWindowIcon(icon)
        Settings.setStyleSheet("")

        def closeEvent(event):
            # Приховати вікно замість закриття
            event.ignore()
            Settings.hide()

        # Пов'язати функцію з вікном
        Settings.closeEvent = closeEvent

        self.shortcut = QtWidgets.QKeySequenceEdit(Settings)
        self.shortcut.setGeometry(QtCore.QRect(20, 40, 261, 51))
        self.shortcut.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.shortcut.setKeySequence(self.hotkey)
        self.shortcut.keySequenceChanged.connect(self.handle_shortcut_changed)
        self.shortcut.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";")
        self.shortcut.setObjectName("shortcut")


        self.Select_path = QtWidgets.QPushButton(Settings)
        self.Select_path.setGeometry(QtCore.QRect(20, 200, 211, 31))
        self.Select_path.setStyleSheet("QPushButton {\n"
"    background-color: #F0F0F0;\n"
"    border: 1px solid #C2C2C2;\n"
"    border-radius: 2px;\n"
"    color: #000000;\n"
"    font-size: 12px;\n"
"    font-weight: bold;\n"
"    min-width: 80px;\n"
"    padding: 6px;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #E0E0E0;\n"
"    border: 1px solid #C2C2C2;\n"
"    color: #000000;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #D3D3D3;\n"
"    border: 1px solid #C2C2C2;\n"
"    color: #000000;\n"
"}\n"
"")
        self.Select_path.setObjectName("Select_path")
        self.Save_settings = QtWidgets.QPushButton(Settings)
        self.Save_settings.setGeometry(QtCore.QRect(20, 380, 211, 31))
        self.Save_settings.setStyleSheet("QPushButton {\n"
"    background-color: #F0F0F0;\n"
"    border: 1px solid #C2C2C2;\n"
"    border-radius: 2px;\n"
"    color: #000000;\n"
"    font-size: 12px;\n"
"    font-weight: bold;\n"
"    min-width: 80px;\n"
"    padding: 6px;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #E0E0E0;\n"
"    border: 1px solid #C2C2C2;\n"
"    color: #000000;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #D3D3D3;\n"
"    border: 1px solid #C2C2C2;\n"
"    color: #000000;\n"
"}\n"
"\n"
"")
        self.Save_settings.setObjectName("Save_settings")


        self.path = QtWidgets.QLineEdit(Settings)
        reg_ex = QRegExp("^[a-zA-Z]:((\/|\\\\)(?:[^\/\\\\\n]+(\/|\\\\))*[^\/\\\\\n]*)?$")
        validator = QRegExpValidator(reg_ex, self.path)
        self.path.setValidator(validator)
        self.path.setGeometry(QtCore.QRect(20, 140, 351, 31))
        self.path.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.path.setStyleSheet("QLineEdit {\n"
"    font: 12pt \"MS Shell Dlg 2\";\n"
"    border: 1px solid #C2C7CB;\n"
"    border-radius: 2px;\n"
"    padding: 4px 8px;\n"
"    font-size: 12px;\n"
"    color: #1E1E1E;\n"
"}\n"
"\n"
"QLineEdit:focus {\n"
"    border: 1px solid #0078D7;\n"
"    outline: none;\n"
"}\n"
"")
        self.path.setText(self.output_path.replace("\\", "/"))
        self.path.setObjectName("path")

        self.label = QtWidgets.QLabel(Settings)
        self.label.setGeometry(QtCore.QRect(20, 9, 251, 31))
        self.label.setStyleSheet("font: 11pt \"MS Shell Dlg 2\";")
        self.label.setObjectName("label")

        self.Save_to = QtWidgets.QLabel(Settings)
        self.Save_to.setGeometry(QtCore.QRect(20, 90, 361, 41))
        self.Save_to.setStyleSheet("font: 11pt \"MS Shell Dlg 2\";")
        self.Save_to.setObjectName("Save_to")

        self.auto_open_pdf = QtWidgets.QCheckBox(Settings)
        with open(self.settings_file) as f:
                settings = json.load(f)
                self.auto_open_pdf_var = settings.get('auto_open_pdf', False)
        self.auto_open_pdf.setChecked(self.auto_open_pdf_var)
        self.auto_open_pdf.setGeometry(QtCore.QRect(20, 240, 191, 51))
        self.auto_open_pdf.setStyleSheet("QCheckBox {\n"
"font-size: 12px;\n"
"}\n"
"\n"
"QCheckBox::indicator {\n"
"width: 18px;\n"
"height: 18px;\n"
"}\n"
"\n"
"QCheckBox::indicator:unchecked {\n"
"border: 2px solid #b0b0b0;\n"
"border-radius: 3px;\n"
"}\n"
"\n"
"QCheckBox::indicator:checked {\n"
"border: 2px solid #0078d7;\n"
"border-radius: 3px;\n"
"background-color: #0078d7;\n"
"}\n"
"\n"
"QCheckBox::indicator:checked:disabled {\n"
"border: 2px solid #b0b0b0;\n"
"border-radius: 3px;\n"
"background-color: #b0b0b0;\n"
"}\n"
"\n"
"QCheckBox::indicator:disabled {\n"
"border: 2px solid #b0b0b0;\n"
"border-radius: 3px;\n"
"background-color: #f0f0f0;\n"
"}\n"
"\n"
"QCheckBox::indicator:checked:focus,\n"
"QCheckBox::indicator:unchecked:focus {\n"
"outline: none;\n"
"}\n"
"\n"
"QCheckBox::indicator:checked:pressed,\n"
"QCheckBox::indicator:unchecked:pressed {\n"
"border: 2px solid #0078d7;\n"
"}\n"
"QCheckBox::indicator:checked:pressed:disabled,\n"
"QCheckBox::indicator:unchecked:pressed:disabled {\n"
"border: 2px solid #b0b0b0;\n"
"}\n"
"QCheckBox::indicator:checked:pressed:focus,\n"
"QCheckBox::indicator:unchecked:pressed:focus {\n"
"outline: none;\n"
"}\n"
"\n"
"QCheckBox::indicator:checked:disabled:focus,\n"
"QCheckBox::indicator:unchecked:disabled:focus {\n"
"outline: none;\n"
"}\n"
"\n"
"QCheckBox::indicator:checked:hover,\n"
"QCheckBox::indicator:unchecked:hover {\n"
"border: 2px solid #0078d7;\n"
"}")
        self.auto_open_pdf.setObjectName("auto_open_pdf")
        self.lineEdit = QtWidgets.QLineEdit(Settings)
        self.lineEdit.setGeometry(QtCore.QRect(20, 330, 351, 31))
        self.lineEdit.setStyleSheet("QLineEdit {\n"
"    font: 12pt \"MS Shell Dlg 2\";\n"
"    border: 1px solid #C2C7CB;\n"
"    border-radius: 2px;\n"
"    padding: 4px 8px;\n"
"    font-size: 12px;\n"
"    color: #1E1E1E;\n"
"}\n"
"\n"
"QLineEdit:focus {\n"
"    border: 1px solid #0078D7;\n"
"    outline: none;\n"
"}\n"
"")
        license_key = "app_id.txt"
        with open(license_key, "r") as file:
            license_key = file.read().strip()
        self.lineEdit.setText(license_key)
        self.lineEdit.setReadOnly(True)
        self.lineEdit.setObjectName("lineEdit")

        self.label_2 = QtWidgets.QLabel(Settings)
        self.label_2.setGeometry(QtCore.QRect(20, 280, 421, 41))
        self.label_2.setStyleSheet("font: 11pt \"MS Shell Dlg 2\";")
        self.label_2.setObjectName("label_2")

        self.retranslateUi(Settings)
        QtCore.QMetaObject.connectSlotsByName(Settings)

        self.Save_settings.clicked.connect(self.save_settings)
        self.Select_path.clicked.connect(self.choose_directory)

    def retranslateUi(self, Settings):
        _translate = QtCore.QCoreApplication.translate
        Settings.setWindowTitle(_translate("Settings", "Ustawienia"))
        self.Select_path.setText(_translate("Settings", "wybierz folder docelowy"))
        self.Save_settings.setText(_translate("Settings", "Zapisz zmiany"))
        self.label.setText(_translate("Settings", "Skrót klawiszowy:"))
        self.Save_to.setText(_translate("Settings", "Zapisz do:"))
        self.auto_open_pdf.setText(_translate("Settings", "Otwórz pdf po pobraniu"))
        self.label_2.setText(_translate("Settings", "Klucz licencyjny:"))

    def handle_shortcut_changed(self, key_sequence):
        if len(key_sequence) > 1:
            self.shortcut.setKeySequence(key_sequence[1])


    def choose_directory(self): 
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(None, "KW pdf - wybierz folder")
        if dir_path:
                self.output_path = dir_path
                self.path.setText(dir_path)

    def save_settings(self):
        hotkey = self.shortcut.keySequence().toString(QtGui.QKeySequence.NativeText)
        path = self.path.text().replace("/", "\\")
        auto_open_pdf = self.auto_open_pdf.isChecked()
        settings = {'hotkey': hotkey, 'output_path': path, 'auto_open_pdf': auto_open_pdf}
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f, indent=4)
        self.check_path()
        self.close_window()

    def close_window(self):
        self.hide()


    def check_path(self):
        # Open the settings file
        with open('settings.json') as f:
            settings = json.load(f)
            output_path = settings.get('output_path', '')
            
            # Check if Documents folder exists, if not set path to Desktop
            documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
            if not os.path.exists(documents_path):
                documents_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            # Check if the output path is empty
            
            if not output_path:
                # If the output path is empty, set it to the Documents folder
                output_path = documents_path
                settings['output_path'] = output_path
                with open('settings.json', 'w') as f:
                    json.dump(settings, f, indent=4)
            
            # Check if the output path exists
            if not os.path.exists(output_path):
                try:
                    # Try to create the output path
                    os.makedirs(output_path)
                except OSError:
                    # If the output path could not be created, set it to the Documents folder
                    output_path = documents_path
                    settings['output_path'] = output_path
                    with open('settings.json', 'w') as f:
                        json.dump(settings, f, indent=4)


    def load_settings(self):
        try:
            with open(self.settings_file) as f:
                settings = json.load(f)
                self.hotkey = settings.get('hotkey')
                self.output_path = settings.get('output_path')
                self.auto_open_pdf = settings.get('auto_open_pdf')
        except FileNotFoundError:
            self.hotkey = 'ctrl+q'
            self.output_path = os.path.join(os.path.expanduser("~"), "Documents")
            if not os.path.exists(self.output_path):
                try:
                    os.makedirs(self.output_path)
                except OSError as e:
                    self.output_path = os.path.join(os.path.expanduser("~"), "Desktop")
            self.auto_open_pdf = False
            self.save_settings()
     
class MainWindow(QMainWindow):

    confirmation_signal = pyqtSignal()

    kw_temp = ""

    def __init__(self):
        super().__init__()
        self.settings_file = 'settings.json'
        self.load_settings()
        self.read_or_generate_app_id()
        self.confirm_window_shown = None

        keyboard.add_hotkey(self.hotkey, self.on_hotkey_press)

    def show_confirmation_signal(self):
        self.confirmation_signal.emit()

    def read_or_generate_app_id(self):
        file_path = os.path.join(os.path.dirname(__file__), 'app_id.txt')

        # Check if app_id.txt file already exists
        if os.path.exists(file_path):
            # If file exists, read existing app ID from file
            with open(file_path, 'r') as f:
                app_id = f.read().strip()
        else:
            # If file does not exist, generate new app ID based on MAC address
            mac_bytes = uuid.getnode().to_bytes(6, byteorder='big')
            mac_string = ':'.join(format(b, '02x') for b in mac_bytes)
            app_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, mac_string))
            # Save app ID to file
            with open(file_path, 'w') as f:
                f.write(str(app_id))
            
            # Set the file as hidden
            if os.name == 'nt':  # For Windows
                os.system(f'attrib +h "{file_path}"')

        os.chmod(file_path, 0o444)

        return app_id
    
    def save_settings(self):
        settings = {'hotkey': self.hotkey, 'output_path': self.output_path, 'auto_open_pdf': self.auto_open_pdf}
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f, indent=4)

    def load_settings(self):
        try:
            with open(self.settings_file) as f:
                settings = json.load(f)
                self.hotkey = settings.get('hotkey')
                self.output_path = settings.get('output_path')
                self.auto_open_pdf = settings.get('auto_open_pdf')
        except FileNotFoundError:
            self.hotkey = 'ctrl+q'
            self.output_path = os.path.join(os.path.expanduser("~"), "Documents")
            if not os.path.exists(self.output_path):
                try:
                    os.makedirs(self.output_path)
                except OSError as e:
                    self.output_path = os.path.join(os.path.expanduser("~"), "Desktop")
            self.auto_open_pdf = False
            self.save_settings()


    def check_license(self, app_id):
        try:
            # Send GET request to API with predefined URL address
            url = "https://kwpdf.pl/API//" + app_id
            response = requests.get(url, verify=False)
            if response.status_code == 200:
                # License is valid, continue with KW report download request
                return True

            elif response.status_code == 402:
                # Create a warning window
                warning_window = tk.Tk()
                warning_window.iconbitmap("ico.ico")
                warning_window.title("Błąd licencji")

                # Set the window size (width x height)
                window_width = 400
                window_height = 150
                window_size = f"{window_width}x{window_height}"
                warning_window.geometry(window_size)

                
                # Center the window on the screen
                screen_width = warning_window.winfo_screenwidth()
                screen_height = warning_window.winfo_screenheight()
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2
                warning_window.geometry(f"+{x}+{y}")

                # Disable window resizing
                warning_window.resizable(False, False)

                # Extract the message and link from the response
                response_data = response.json()
                message = response_data["msg"]
                link = response_data["link"]

                # Define the click event for opening the link
                def open_link(event):
                    webbrowser.open(link)

                # Create a label with the message
                message_label = tk.Label(warning_window, text=message, wraplength=380)  # Додано параметр wraplength
                message_label.pack(padx=10, pady=10)

                # Create a clickable link label
                link_label = tk.Label(warning_window, text=link, fg="blue", cursor="hand2")
                link_label.pack(padx=20, pady=10)
                link_label.bind("<Button-1>", open_link)

                def ok_pressed():
                    webbrowser.open(link)
                    warning_window.destroy()

                def cancel_pressed():
                    warning_window.destroy()

                # Create the Cancel button
                cancel_button_style = ttk.Style()
                cancel_button_style.configure("Cancel.TButton", font=("Segoe UI", 10, "bold"))

                cancel_button = ttk.Button(warning_window, text="Cancel", style="OK.TButton", command=cancel_pressed)
                cancel_button.pack(side=tk.RIGHT, padx=10)

                # Create the OK button
                ok_button_style = ttk.Style()
                ok_button_style.configure("OK.TButton", font=("Segoe UI", 10, "bold"))

                ok_button = ttk.Button(warning_window, text="OK", style="OK.TButton", command=ok_pressed)
                ok_button.pack(side=tk.RIGHT, padx=10)

                winsound.MessageBeep()

                warning_window.attributes('-topmost', True)
                warning_window.mainloop()
                #ctypes.windll.user32.MessageBoxW(None, response.text, "Błąd licencji", MB_ICONHAND | MB_OK | 0x1000)
                return False
            else:
                # Handle other response codes here
                ctypes.windll.user32.MessageBoxW(None, response.status_code, "Unexpected response code:", MB_ICONHAND | MB_OK | 0x1000)
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(None, str(e), "Błąd licencji", MB_ICONHAND | MB_OK | 0x1000)
            return False

    def calculate_kw_control_digit(self, kw):
        symbols = {
            '0': 0, '1': 1, '2': 2, '3': 3, '4': 4,
            '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
            'A': 11, 'B': 12, 'C': 13, 'D': 14, 'E': 15,
            'F': 16, 'G': 17, 'H': 18, 'I': 19, 'J': 20,
            'K': 21, 'L': 22, 'M': 23, 'N': 24, 'O': 25,
            'P': 26, 'R': 27, 'S': 28, 'T': 29,
            'U': 30, 'W': 31, 'X': 10, 'Y': 32, 'Z': 33
        }
        weights = [1, 3, 7, 1, 3, 7, 1, 3, 7, 1, 3, 7]
        t = kw.replace('/', '') #remove the slash from the copied key
        w = t[:-1] #delete the wrong control number
        sum = 0
        for i in range(len(w)):
            try:
                sum += symbols[w[i]] * weights[i]
            except KeyError:
                return None # if symbol is not found in the dictionary, return None
                
        control_digit = sum % 10
        return control_digit
    

    def on_hotkey_press(self):
        app_id = self.read_or_generate_app_id()
        if self.check_license(app_id):
            try:
                selected_key = pyperclip.paste().strip().replace(' ', '')
                if len(selected_key) < 15:
                    return ctypes.windll.user32.MessageBoxW(None, "Nie znaleziono numeru księgi wieczystej w zaznaczonym tekście", "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
                
                match = re.search('([A-Z0-9]{4}[/][0-9]{8}[/][0-9])', selected_key)
                if not match:
                    return ctypes.windll.user32.MessageBoxW(None, "Nie znaleziono numeru księgi wieczystej w zaznaczonym tekście", "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
                else:
                    kw = match.group(0)
                    # Check the control digit
                    control_digit = int(kw[-1])
                    calculated_control_digit = self.calculate_kw_control_digit(kw)
                    if control_digit != calculated_control_digit:
                        # Display correct KW-Key
                        kw2 = kw[:-1] + str(calculated_control_digit)
                        return ctypes.windll.user32.MessageBoxW(None, f'Zaznaczony numer księgi wieczystej: {kw}\nma nieprawidłową cyfrę kontrolną. Prawidłowy numer księgi wieczystej: {kw2}', "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
                            # If we need we can Replace the incorrect control digit with the correct one
                            # kw = kw[:-1] + str(calculated_control_digit)
                MainWindow.kw_temp = kw
                self.show_confirmation_signal()
            except Exception as e:
                return ctypes.windll.user32.MessageBoxW(None, str(e), "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
               

    def show_confirmation(self):
        self.confirm_window_shown = ConfirmWindow()
        self.confirm_window_shown.setupUi(self.confirm_window_shown)
        self.confirm_window_shown.setWindowModality(QtCore.Qt.WindowModal)
        self.confirm_window_shown.setWindowState(QtCore.Qt.WindowActive)
        self.confirm_window_shown.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        self.confirm_window_shown.show()


class ConfirmWindow(QWidget):

    lock = Lock()

    def __init__(self):
        super().__init__()
        self.settings_file = 'settings.json'
        self.load_settings()
        self.read_or_generate_app_id()
        self.check_path()

    def setupUi(self, ask_confirmation):
        self.confirm_window = ask_confirmation
        ask_confirmation.setObjectName("ask_confirmation")
        ask_confirmation.setWindowFlags(ask_confirmation.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        ask_confirmation.resize(500, 200)
        ask_confirmation.setMinimumSize(QtCore.QSize(500, 200))
        ask_confirmation.setMaximumSize(QtCore.QSize(500, 200))
        ask_confirmation.setBaseSize(QtCore.QSize(0, 0))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("1.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        ask_confirmation.setWindowIcon(icon)


        # Додати функцію для закриття вікна
        def closeEvent(event):
            # Приховати вікно замість закриття
            event.ignore()
            ask_confirmation.hide()

        # Пов'язати функцію з вікном
        ask_confirmation.closeEvent = closeEvent


        self.path = QtWidgets.QLineEdit(ask_confirmation)
        # Create a regular expression that defines what characters can be entered
        reg_ex = QRegExp("^[a-zA-Z]:((\/|\\\\)(?:[^\/\\\\\n]+(\/|\\\\))*[^\/\\\\\n]*)?$")
        validator = QRegExpValidator(reg_ex, self.path)
        self.path.setValidator(validator)
        self.path.setGeometry(QtCore.QRect(40, 40, 421, 31))
        self.path.setStyleSheet("QLineEdit {\n"
"    font: 12pt \"MS Shell Dlg 2\";\n"
"    border: 1px solid #C2C7CB;\n"
"    border-radius: 2px;\n"
"    padding: 4px 8px;\n"
"    font-size: 12px;\n"
"    color: #1E1E1E;\n"
"}\n"
"\n"
"QLineEdit:focus {\n"
"    border: 1px solid #0078D7;\n"
"    outline: none;\n"
"}\n"
"")
        self.path.setText(self.output_path.replace("\\", "/"))
        self.path.setObjectName("path")


        self.select_path = QtWidgets.QPushButton(ask_confirmation)
        self.select_path.setGeometry(QtCore.QRect(130, 90, 201, 31))
        self.select_path.setStyleSheet("QPushButton {\n"
"    background-color: #F0F0F0;\n"
"    border: 1px solid #C2C2C2;\n"
"    border-radius: 2px;\n"
"    color: #000000;\n"
"    font-size: 14px;\n"
"    font-weight: bold;\n"
"    min-width: 80px;\n"
"    padding: 6px;\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"    background-color: #E0E0E0;\n"
"    border: 1px solid #C2C2C2;\n"
"    color: #000000;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: #D3D3D3;\n"
"    border: 1px solid #C2C2C2;\n"
"    color: #000000;\n"
"}\n"
"")
        self.select_path.setObjectName("select_path")


        self.Save_to = QtWidgets.QLabel(ask_confirmation)
        self.Save_to.setGeometry(QtCore.QRect(190, 10, 281, 31))
        self.Save_to.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";")
        self.Save_to.setObjectName("Save_to")

        self.buttonBox = QtWidgets.QDialogButtonBox(ask_confirmation)
        self.buttonBox.setGeometry(QtCore.QRect(110, 130, 201, 61))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.buttonBox.sizePolicy().hasHeightForWidth())
        self.buttonBox.setSizePolicy(sizePolicy)
        self.buttonBox.setMinimumSize(QtCore.QSize(0, 0))
        self.buttonBox.setBaseSize(QtCore.QSize(0, 0))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.buttonBox.setFont(font)
        self.buttonBox.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.buttonBox.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.buttonBox.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates))
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")


        self.retranslateUi(ask_confirmation)
        QtCore.QMetaObject.connectSlotsByName(ask_confirmation)

        self.select_path.clicked.connect(self.choose_directory)
        self.buttonBox.accepted.connect(self.handle_ok_pressed)
        self.buttonBox.rejected.connect(self.handle_cancel_pressed)

    def retranslateUi(self, ask_confirmation):
        _translate = QtCore.QCoreApplication.translate
        ask_confirmation.setWindowTitle(_translate("ask_confirmation", "KW pdf - wybierz folder"))
        self.select_path.setText(_translate("ask_confirmation", "wybierz folder docelowy"))
        self.Save_to.setText(_translate("ask_confirmation", "Zapisz do:"))


    def handle_ok_pressed(self):
        self.save_path()
        self.close_window()
        self.start_get_pdf()
        

    def handle_cancel_pressed(self):
        self.close_window()
        MainWindow.kw_temp = ""


    def choose_directory(self): 
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(None, "KW pdf - wybierz folder")
        if dir_path:
                self.output_path = dir_path
                self.path.setText(dir_path)

    
    def close_window(self):
        self.close()

    def save_path(self):
        path = self.path.text().replace("/", "\\")
        with open(self.settings_file, 'r') as f:
            settings = json.load(f)
            settings['output_path'] = path
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f, indent=4)
        self.check_path()

    def save__file_settings(self):
        settings = {'hotkey': self.hotkey, 'output_path': self.output_path, 'auto_open_pdf': self.auto_open_pdf}
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f, indent=4)


    def check_path(self):
        # Open the settings file
        with open('settings.json') as f:
            settings = json.load(f)
            output_path = settings.get('output_path', '')
            
            # Check if Documents folder exists, if not set path to Desktop
            documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
            if not os.path.exists(documents_path):
                documents_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            # Check if the output path is empty
            
            if not output_path:
                # If the output path is empty, set it to the Documents folder
                output_path = documents_path
                settings['output_path'] = output_path
                with open('settings.json', 'w') as f:
                    json.dump(settings, f, indent=4)
            
            # Check if the output path exists
            if not os.path.exists(output_path):
                try:
                    # Try to create the output path
                    os.makedirs(output_path)
                    self.output_path = output_path
                except OSError:
                    # If the output path could not be created, set it to the Documents folder
                    output_path = documents_path
                    settings['output_path'] = output_path
                    with open('settings.json', 'w') as f:
                        json.dump(settings, f, indent=4)


    def load_settings(self):
        try:
            with open(self.settings_file) as f:
                settings = json.load(f)
                self.hotkey = settings.get('hotkey')
                self.output_path = settings.get('output_path')
                self.auto_open_pdf = settings.get('auto_open_pdf')
        except FileNotFoundError:
            self.hotkey = 'ctrl+q'
            self.output_path = os.path.join(os.path.expanduser("~"), "Documents")
            if not os.path.exists(self.output_path):
                try:
                    os.makedirs(self.output_path)
                except OSError as e:
                    self.output_path = os.path.join(os.path.expanduser("~"), "Desktop")
            self.auto_open_pdf = False
            self.save__file_settings()


    def read_or_generate_app_id(self):
        file_path = os.path.join(os.path.dirname(__file__), 'app_id.txt')

        # Check if app_id.txt file already exists
        if os.path.exists(file_path):
            # If file exists, read existing app ID from file
            with open(file_path, 'r') as f:
                app_id = f.read().strip()
        else:
            # If file does not exist, generate new app ID based on MAC address
            mac_bytes = uuid.getnode().to_bytes(6, byteorder='big')
            mac_string = ':'.join(format(b, '02x') for b in mac_bytes)
            app_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, mac_string))
            # Save app ID to file
            with open(file_path, 'w') as f:
                f.write(str(app_id))
            
            # Set the file as hidden
            if os.name == 'nt':  # For Windows
                os.system(f'attrib +h "{file_path}"')
            
        os.chmod(file_path, 0o444)

        return app_id
    
    def start_get_pdf(self):
        kw = MainWindow.kw_temp
        thread = Thread(target=self.get_pdf, args=(kw,))
        thread.start()

        
    def get_pdf(self, key):
        with self.lock:
            MAIN_DIR = os.path.dirname(os.path.abspath(__file__))
            executable_path = os.path.join(MAIN_DIR, 'chromedriver.exe')

            # Set up the Chrome driver with specified options
            options = webdriver.ChromeOptions()
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument("--window-position=-10000,-10000")
            options.add_argument("--disable-extensions")
            options.add_experimental_option('excludeSwitches', ['enable-automation'])
            options.add_experimental_option('useAutomationExtension', False)

            # Create a new service for the Chrome driver without cmd
            s = Service(executable_path)
            flag = 0x08000000
            s.creation_flags = flag

            webdriver.common.service.subprocess.Popen = functools.partial(webdriver.common.service.subprocess.Popen, creationflags=flag)

            # Set up the Chrome driver with specified options and service
            driver = webdriver.Chrome(service=s, options=options)
            driver.minimize_window()

            
            try:                   
                key_parts = key.split('/')

                driver.implicitly_wait(10)

                # Go to the website URL
                driver.get("https://ekw.ms.gov.pl/eukw_ogol/menu.do")

                # Wait for the page to load
                driver.implicitly_wait(10)

                # Find the link you want to click and click it
                xpath = '/html/body/div/div[2]/div/div/div[2]/div[1]/ul/li[1]/span[1]/a'
                driver.find_element(By.XPATH, xpath).click()

                driver.implicitly_wait(10)


                cell1 = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/form/div[1]/div[2]/div[2]/div[1]/span[1]/input")
                cell2 = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/form/div[1]/div[2]/div[2]/div[1]/span[3]/input")
                cell3 = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/form/div[1]/div[2]/div[2]/div[1]/span[5]/input")


                cell1.send_keys(key_parts[0])
                cell2.send_keys(key_parts[1])
                cell3.send_keys(key_parts[2])
    


                # Wait for the page to load
                driver.implicitly_wait(10)

                # Click the WYSZUKAJ KSIĘGI button
                search_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/form/div[2]/div[2]/button[2]")
                search_button.click()

                # Wait for the page to load
                driver.implicitly_wait(10)

                # Click the PRZEGLĄDANIE AKTUALNEJ TREŚCI KW button
                content_button = driver.find_element(By.XPATH, "/html/body/div/div[2]/div/div[4]/form/div[1]/button[1]")
                content_button.click()

                driver.implicitly_wait(10)

                # find the table with the identifier "nawigacja"
                nawigacja_table = driver.find_element(By.ID, 'nawigacja')

                # find all forms in the table
                forms = nawigacja_table.find_elements(By.TAG_NAME, 'form')

                # get a list of values of the "value" attribute for each tab
                tab_values = []
                for form in forms:
                    tab = form.find_element(By.CSS_SELECTOR, 'input[type="submit"]')
                    tab_value = tab.get_attribute('value')
                    tab_values.append(tab_value)


                htmls = []
                htmls.append(driver.find_element(By.XPATH, "/html/body/h2").get_attribute('outerHTML'))
                htmls.append(driver.find_element(By.XPATH, "/html/body/h4").get_attribute('outerHTML'))
                htmls.append(driver.find_element(By.XPATH, "/html/body/h3").get_attribute('outerHTML'))
                for tab_value in tab_values:
                    # Click on the tab
                    tab_button = driver.find_element(By.XPATH, f"//input[@value='{tab_value}']")
                    tab_button.click()
                    driver.implicitly_wait(10)

                    # Get the HTML code of the current page
                    html = driver.page_source
                    match = re.search("<div id=\"contentDzialu\">[\s\S]*</div>", html, re.MULTILINE)
                    if match:
                        htmls.append(html[match.start():match.end()])

                # Combine all the HTML pages into one string
                combined_html = ''.join(htmls)

                driver.quit()

                # Define CSS styles for the PDF

                styles = """
                    @page {
                        size: A4 portrait;
                        margin: 5px 5px 5px 5px;
                    }
                    body {
                        font-size: 10pt;
                    }

                    h2, h3, h4 {
                        color: black !important;
                        display: block;
                        font-weight: normal;
                        text-align: center;
                    }

                    h4 {
                        font-size: 11pt;
                    }

                    table {
                        table-layout: fixed;
                        width: 100%;
                        white-space: nowrap;
                        border-collapse: collapse;
                    }

                    th, td {
                        padding: 5px;
                        border: 1px solid black;
                        white-space: normal;
                        overflow-wrap: auto
                    }

                    td.csNDBDane {
                        border-bottom: none;

                    }

                    td.csDane {
                        border-top: none;
                    }

                    .csTTytul {
                        background-color: white;
                        text-align: center;
                        font-size: 18px;
                        font-weight: bold;
                        border: 1px solid black;
                        font-family: Verdana;
                        padding: 1px;
                    }
                    .csPodTytulClean {
                        background-color: white;
                        text-align: center;
                        font-size: 16px;
                        font-weight: bold;
                        border: 1px solid black;
                        font-family: Verdana;
                    }
                """

                css = CSS(string=styles)

                # Generate PDF with options
                html = HTML(string=combined_html)
                now = datetime.now()
                date_time = now.strftime("%Y-%m-%d_%H-%M")
                

                with open(self.settings_file) as f:
                    settings = json.load(f)
                
                self.check_path()

                output_file = os.path.join(self.output_path, f'KW {key_parts[0]}_{key_parts[1]}_{key_parts[2]} {date_time}.pdf')

                html.write_pdf(output_file, stylesheets=[css])

                with open('settings.json') as f:
                    settings = json.load(f)
                    if settings['auto_open_pdf']:
                        os.startfile(output_file)


            # Display an error in the Windows message box
            except TimeoutException as timeout_error:
                driver.quit()
                return ctypes.windll.user32.MessageBoxW(None, "Limit czasu oczekiwania został przekroczony. strona jest nieaktywna. Spróbuj ponownie za chwilę", "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
            
            except RequestException:
                driver.quit()
                return ctypes.windll.user32.MessageBoxW(None, "Unable to connect to server. Please check your internet connection and try again.", "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
            
            except NoSuchElementException as kw_error:
                driver.quit()
                return ctypes.windll.user32.MessageBoxW(None, f'Księga wieczysta numer {key_parts[0]}/{key_parts[1]}/{key_parts[2]} nie istnieje', 'KW pdf - błąd', MB_ICONHAND | MB_OK | 0x1000)

            except NoSuchWindowException as NoSuchWindow:
                driver.quit()
                return ctypes.windll.user32.MessageBoxW(None, "Ktoś zamknął okno ", "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
            
            except WebDriverException:
                driver.quit()
                return ctypes.windll.user32.MessageBoxW(None, "Błąd połączenia z internetem lub aplikacji. Skontaktuj się z administratorem aplikacji", "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)

            except Exception as e:
                driver.quit()
                return ctypes.windll.user32.MessageBoxW(None, str(e), "KW pdf - błąd", MB_ICONHAND | MB_OK | 0x1000)
            
            finally:
                time.sleep(1)


if __name__ == '__main__':
    app = QApplication([])
    main_window = MainWindow()
    confirm_window = ConfirmWindow()

    system_tray_icon = SystemTrayIcon()
    system_tray_icon.show()

    main_window.confirmation_signal.connect(main_window.show_confirmation)

    app.exec_()