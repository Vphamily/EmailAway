from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox
from PyQt6.QtCore import pyqtSignal
import hashlib
import json
from datetime import datetime

class LoginDialog(QDialog):
    authenticated = pyqtSignal(dict)

    def __init__(self, license_data, user_data):
        super().__init__()
        self.license_data = license_data
        self.user_data = user_data
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        self.setWindowTitle("Login")
        self.resize(300, 200)

        # Username
        layout.addWidget(QLabel("Username:"))
        self.username_edit = QLineEdit(self)
        layout.addWidget(self.username_edit)

        # Password
        layout.addWidget(QLabel("Password:"))
        self.password_edit = QLineEdit(self)
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_edit)

        # Login Button
        login_button = QPushButton("Login", self)
        login_button.clicked.connect(self.check_credentials)
        layout.addWidget(login_button)

    def check_credentials(self):
        username = self.username_edit.text()
        password = self.password_edit.text()
        hashed_password = hashlib.sha256(password.encode()).hexdigest()

        # Verify user credentials
        user = next((user for user in self.user_data if
                     user['username'] == username and user['password_hash'] == hashed_password), None)
        if user:
            # Check if the user's account is active
            if not user.get('active', True):  # Default to False if 'active' key is missing
                QMessageBox.warning(self, "Login Failed", "Your account is not active.")
                return

            license_key = user['license_key']
            license = next((lic for lic in self.license_data if lic['license_key'] == license_key), None)
            if not license:
                QMessageBox.warning(self, "Login Failed", "No license found for this user.")
            elif not license['active']:
                QMessageBox.warning(self, "Login Failed", "Your license is not active.")
            elif not license['subscription_paid']:
                QMessageBox.warning(self, "Login Failed", "Your subscription has not been paid.")
            elif datetime.now() >= datetime.fromisoformat(license['expiration_date']):
                QMessageBox.warning(self, "Login Failed", "Your license has expired.")
            else:
                QMessageBox.information(self, "Login Successful", "You have successfully logged in.")
                self.authenticated.emit(user)
                self.accept()
        else:
            QMessageBox.critical(self, "Login Failed", "Invalid username or password.")

