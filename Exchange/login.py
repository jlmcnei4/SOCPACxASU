from PyQt5.QtCore import Qt
from PyQt5.QtGui import *
from exchangelib import Credentials, Account, DistributionList, Configuration
from exchangelib.indexed_properties import EmailAddress
from exchangelib import Build, NTLM
from PyQt5.QtWidgets import *

import sys


class Creds:
    def __init__(self, username, password):
        self.u = username
        self.p = password


class AnotherWindow(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.label = QLabel("Another Window")
        layout.addWidget(self.label)
        self.setLayout(layout)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.w = None
        self.setWindowTitle("Aviation Management Dashboard")
        self.window = QWidget()

        layout = QVBoxLayout()
        self.label = QLabel('Please enter Outlook email address and password to log in.')
        self.font = QFont()
        self.font.setPointSize(16)
        self.label.setFont(self.font)
        layout.addWidget(self.label)

        self.username = QLineEdit()
        self.username.setFont(self.font)
        self.password = QLineEdit()
        self.password.setFont(self.font)
        self.password.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.username)
        layout.addWidget(self.password)

        self.button = QPushButton('Log In')
        self.button.setFont(self.font)
        self.button.clicked.connect(self.on_button_clicked)
        layout.addWidget(self.button)

        self.window.setLayout(layout)
        self.setCentralWidget(self.window)
        self.setContentsMargins(250, 250, 250, 250)

    def on_button_clicked(self, checked):
        global userCreds, account
        userCreds.u = self.username.text()
        userCreds.p = self.password.text()
        self.username.clear()
        self.password.clear()
        success = updateCredentials(userCreds.u, userCreds.p)
        if success == 0:
            if self.w is None:
                self.w = AnotherWindow()
            self.w.show()

        else:
            alert = QMessageBox()
            alert.setText('Invalid credentials entered, \n Please try again.')
            alert.exec()


def updateCredentials(username, password):
    global credentials, account, folder, config
    try:
        credentials = Credentials(username, password)
        config = Configuration(server='outlook.office365.com', credentials=credentials)
        account = Account(userCreds.u, credentials=credentials, config=config, autodiscover=False)
        folder = account.root / 'AllContacts'
        return 0
    except ValueError:
        return -1


userCreds = Creds('', '')
config = None
credentials = Credentials(userCreds.u, userCreds.p)
account = None
folder = None
app = QApplication(sys.argv)
app.setStyle('Fusion')

dark_palette = QPalette()

dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
dark_palette.setColor(QPalette.WindowText, Qt.white)
dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
dark_palette.setColor(QPalette.ToolTipText, Qt.white)
dark_palette.setColor(QPalette.Text, Qt.white)
dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
dark_palette.setColor(QPalette.ButtonText, Qt.white)
dark_palette.setColor(QPalette.BrightText, Qt.red)
dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
dark_palette.setColor(QPalette.HighlightedText, Qt.black)

app.setPalette(dark_palette)
app.setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }")


w = MainWindow()
w.show()
app.exec()
