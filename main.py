import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import threading
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
                             QTextEdit, QDialog, QCheckBox,
                             QListWidget, QComboBox, QLabel, QFileDialog, QMessageBox, QDateTimeEdit, QTabWidget)
from PyQt6.QtCore import QDate, QDateTime, pyqtSignal
from datetime import datetime, timedelta
import json
import os
import win32com.client as win32  # For Outlook integration
from PyQt6.QtGui import QIcon
from LoginDialog import LoginDialog

class EmailWindow(QDialog):
    email_sent_signal = pyqtSignal(str)

    def __init__(self, parent=None, callback=None, email_data=None):
        super().__init__(parent)
        self.callback = callback
        self.email_data = email_data or {}
        self.attachments = []
        self.schedule_enabled = False
        self.setWindowTitle("Create Your Task")
        self.resize(500, 600)
        self.init_ui()

        self.email_sent_signal.connect(self.handle_email_sent)

    def handle_email_sent(self, message):
        print(message)  # Print the success message to console or handle it appropriately

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Create a tab widget
        self.tabs = QTabWidget(self)

        # Create individual tabs for email, schedule, and server config
        self.email_tab = QWidget()
        self.schedule_tab = QWidget()
        self.server_tab = QWidget()

        # Add the tabs to the QTabWidget
        self.tabs.addTab(self.email_tab, "Email")
        self.tabs.addTab(self.schedule_tab, "Schedule")
        self.tabs.addTab(self.server_tab, "Server Config")

        layout.addWidget(self.tabs)

        # Schedule Toggle Checkbox
        self.schedule_toggle_checkbox = QCheckBox("Enable Schedule", self)
        self.schedule_toggle_checkbox.setChecked(False)  # Default to unchecked
        self.schedule_toggle_checkbox.stateChanged.connect(self.toggle_schedule_tab)
        layout.addWidget(self.schedule_toggle_checkbox)

        # Set up each tab
        self.init_email_tab()
        self.init_schedule_tab()
        self.init_server_tab()

        # Save and Test Buttons
        save_button = QPushButton("Save", self)
        save_button.clicked.connect(self.save_email)
        layout.addWidget(save_button)

        test_send_button = QPushButton("Test Send", self)
        test_send_button.clicked.connect(self.test_send_email)
        layout.addWidget(test_send_button)

        cancel_button = QPushButton("Cancel", self)
        cancel_button.clicked.connect(self.cancel_email)
        layout.addWidget(cancel_button)

        if self.email_data:
            self.load_email_data()

    def init_email_tab(self):
        email_layout = QVBoxLayout(self.email_tab)

        # Task name
        email_layout.addWidget(QLabel("Task Name:"))
        self.task_name_edit = QLineEdit(self)
        email_layout.addWidget(self.task_name_edit)

        # Email Fields
        email_layout.addWidget(QLabel("To:"))
        self.to_edit = QLineEdit(self)
        email_layout.addWidget(self.to_edit)

        email_layout.addWidget(QLabel("Subject:"))
        self.subject_edit = QLineEdit(self)
        email_layout.addWidget(self.subject_edit)

        email_layout.addWidget(QLabel("Message:"))

        # Create a horizontal layout for the message editor
        message_layout = QHBoxLayout()
        self.message_edit = QTextEdit(self)
        message_layout.addWidget(self.message_edit)

        email_layout.addLayout(message_layout)  # Add the horizontal layout to the main layout

        # Attachments
        email_layout.addWidget(QLabel("Attachments:"))
        self.attachment_list = QListWidget(self)
        email_layout.addWidget(self.attachment_list)

        # Create a horizontal layout for Add and Remove buttons
        button_layout = QHBoxLayout()

        add_attachment_button = QPushButton("Add Attachment", self)
        add_attachment_button.clicked.connect(self.add_attachment)
        button_layout.addWidget(add_attachment_button)

        remove_attachment_button = QPushButton("Remove Attachment", self)
        remove_attachment_button.clicked.connect(self.remove_attachment)
        button_layout.addWidget(remove_attachment_button)

        email_layout.addLayout(button_layout)  # Add the button layout to the main layout

        # Insert Date Button
        insert_date_button = QPushButton("{Insert Date}", self)
        insert_date_button.setFixedSize(100, 30)
        insert_date_button.clicked.connect(self.insert_date)

        email_layout.addWidget(insert_date_button)  # Add Insert Date button below the Remove button

    def insert_date(self):
        """Function to insert date variables into the message body."""
        date_options = [
            "{Date} - Current date",
            "{DateTime} - Current date and time",
            "{Day} - Current day",
            "{DayOfWeek} - Current day of the week",
            "{DayOfYear} - Day of the year",
            "{DateInDays(-7)} - Current date minus X days"  # No prompt; user will input X directly
        ]

        # Create a dialog for the user to choose a date variable
        dialog = QDialog(self)
        dialog.setWindowTitle("Insert Date Variable")
        layout = QVBoxLayout(dialog)

        # Add options to the dialog
        for option in date_options:
            button = QPushButton(option)
            button.clicked.connect(lambda checked, option=option: self.insert_variable(option.split(' - ')[0]))
            layout.addWidget(button)

        dialog.setLayout(layout)
        dialog.exec()

    def insert_variable(self, variable):
        """Insert the selected variable into the message body at the current cursor position"""
        cursor = self.message_edit.textCursor()
        cursor.insertText(variable)  # Insert the

    def init_schedule_tab(self):
        schedule_layout = QVBoxLayout(self.schedule_tab)

        schedule_layout.addWidget(QLabel("Frequency:"))
        self.frequency_combo = QComboBox(self)
        self.frequency_combo.addItems(["Run Once", "Daily", "Weekly", "Monthly"])
        schedule_layout.addWidget(self.frequency_combo)

        schedule_layout.addWidget(QLabel("Start Date:"))
        self.start_date_edit = QDateTimeEdit(self)
        self.start_date_edit.setDisplayFormat("MM/dd/yyyy")
        self.start_date_edit.setDate(QDate.currentDate())
        schedule_layout.addWidget(self.start_date_edit)

        schedule_layout.addWidget(QLabel("Time:"))
        time_layout = QHBoxLayout()
        self.time_edit = QLineEdit(self)
        self.am_pm_combo = QComboBox(self)
        self.am_pm_combo.addItems(["AM", "PM"])
        time_layout.addWidget(self.time_edit)
        time_layout.addWidget(self.am_pm_combo)
        schedule_layout.addLayout(time_layout)

    def init_server_tab(self):
        server_layout = QVBoxLayout(self.server_tab)

        server_layout.addWidget(QLabel("Server Type:"))
        self.server_type_combo = QComboBox(self)
        self.server_type_combo.addItems(["SMTP", "Outlook", "Gmail", "Yahoo"])
        self.server_type_combo.currentIndexChanged.connect(self.autofill_smtp_settings)
        server_layout.addWidget(self.server_type_combo)

        server_layout.addWidget(QLabel("Email Address:"))
        self.email_address_edit = QLineEdit(self)
        server_layout.addWidget(self.email_address_edit)

        server_layout.addWidget(QLabel("Password:"))
        self.password_edit = QLineEdit(self)
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        server_layout.addWidget(self.password_edit)

        # Additional label for Gmail and Yahoo app password message
        self.app_password_label = QLabel("", self)
        self.app_password_label.setStyleSheet("color: red")
        server_layout.addWidget(self.app_password_label)

        server_layout.addWidget(QLabel("SMTP Server:"))
        self.smtp_server_edit = QLineEdit(self)
        server_layout.addWidget(self.smtp_server_edit)

        server_layout.addWidget(QLabel("SMTP Port:"))
        self.smtp_port_edit = QLineEdit(self)
        server_layout.addWidget(self.smtp_port_edit)

    def autofill_smtp_settings(self):
        """Auto-fill SMTP settings based on selected server type."""
        selected_server = self.server_type_combo.currentText()

        if selected_server == "Gmail":
            self.smtp_server_edit.setText("smtp.gmail.com")
            self.smtp_port_edit.setText("587")
            self.app_password_label.setText("Please use App password")
            self.smtp_server_edit.setEnabled(True)
            self.smtp_port_edit.setEnabled(True)

        elif selected_server == "Yahoo":
            self.smtp_server_edit.setText("smtp.mail.yahoo.com")
            self.smtp_port_edit.setText("587")
            self.app_password_label.setText("Please use App password")
            self.smtp_server_edit.setEnabled(True)
            self.smtp_port_edit.setEnabled(True)

        elif selected_server == "Outlook":
            self.smtp_server_edit.clear()
            self.smtp_port_edit.clear()
            self.app_password_label.setText("")
            self.smtp_server_edit.setEnabled(False)
            self.smtp_port_edit.setEnabled(False)

        else:
            self.smtp_server_edit.clear()
            self.smtp_port_edit.clear()
            self.app_password_label.setText("")
            self.smtp_server_edit.setEnabled(True)
            self.smtp_port_edit.setEnabled(True)

    def toggle_schedule_tab(self, state):
        """Enable or disable the schedule tab based on the checkbox state"""
        if state == 2:  # Checked
            self.tabs.setTabEnabled(1, True)
            self.schedule_enabled = True
        else:
            self.tabs.setTabEnabled(1, False)
            self.schedule_enabled = False

    def add_attachment(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select File")
        if file_name:
            self.attachments.append(file_name)
            self.attachment_list.addItem(os.path.basename(file_name))

    def remove_attachment(self):
        current_row = self.attachment_list.currentRow()
        if current_row >= 0:
            self.attachments.pop(current_row)
            self.attachment_list.takeItem(current_row)

    def load_email_data(self):
        # Email Tab Data
        self.task_name_edit.setText(self.email_data.get('task_name', ''))
        self.to_edit.setText(self.email_data.get('to', ''))
        self.subject_edit.setText(self.email_data.get('subject', ''))
        self.message_edit.setPlainText(self.email_data.get('message', ''))

        # Load attachments into the attachment list
        self.attachments = self.email_data.get('attachments', [])
        self.attachment_list.clear()
        for attachment in self.attachments:
            self.attachment_list.addItem(os.path.basename(attachment))

        # Schedule Tab Data
        schedule_data = self.email_data.get('schedule', {})
        self.frequency_combo.setCurrentText(schedule_data.get('frequency', 'Run Once'))

        try:
            start_date = datetime.strptime(schedule_data.get('start_date', ''), '%m/%d/%Y')
            self.start_date_edit.setDate(QDate(start_date.year, start_date.month, start_date.day))
        except ValueError:
            self.start_date_edit.setDate(QDate.currentDate())

        time_parts = schedule_data.get('time', '').split()
        if len(time_parts) == 2:
            self.time_edit.setText(time_parts[0])
            self.am_pm_combo.setCurrentText(time_parts[1])

        # Server Tab Data
        server_data = self.email_data.get('server', {})
        self.server_type_combo.setCurrentText(server_data.get('server_type', 'SMTP'))
        self.email_address_edit.setText(server_data.get('email_address', ''))
        self.password_edit.setText(server_data.get('password', ''))
        self.smtp_server_edit.setText(server_data.get('smtp_server', ''))
        self.smtp_port_edit.setText(server_data.get('smtp_port', ''))

    def save_email(self):
        task_name = self.task_name_edit.text()
        if not task_name:
            QMessageBox.critical(self, "Error", "Task Name cannot be empty!")
            return

        # Get data from the tabs
        start_date_str = self.start_date_edit.date().toString("MM/dd/yyyy")
        full_time = f"{self.time_edit.text()} {self.am_pm_combo.currentText()}"

        email_data = {
            'task_name': task_name,
            'to': self.to_edit.text(),
            'subject': self.subject_edit.text(),
            'message': self.message_edit.toPlainText(),
            'schedule': {
                'frequency': self.frequency_combo.currentText(),
                'start_date': start_date_str,
                'time': full_time
            },
            'server': {
                'server_type': self.server_type_combo.currentText(),
                'email_address': self.email_address_edit.text(),
                'password': self.password_edit.text(),
                'smtp_server': self.smtp_server_edit.text(),
                'smtp_port': self.smtp_port_edit.text()
            },
            'attachments': self.attachments,
            'schedule_enabled': self.schedule_enabled
        }

        # Pass the email data back to the main UI via the callback
        self.callback(email_data)

        # Close the email composer window
        self.close()

    def test_send_email(self):
        server_type = self.server_type_combo.currentText()

        if server_type == "Outlook":
            self.send_email_via_outlook()
        else:
            # Prepare email data and send it immediately for testing
            email_data = {
                'to': self.to_edit.text(),
                'subject': self.subject_edit.text(),
                'message': self.message_edit.toPlainText(),
                'server': {
                    'server_type': self.server_type_combo.currentText(),
                    'email_address': self.email_address_edit.text(),
                    'password': self.password_edit.text(),
                    'smtp_server': self.smtp_server_edit.text(),
                    'smtp_port': self.smtp_port_edit.text()
                },
                'attachments': self.attachments,
            }
            self.send_email(email_data)

    def send_email(self, email_data):
        """Generic method to send email based on the server type"""
        server_type = email_data['server']['server_type']

        # Replace date placeholders in the message body
        message_body = self.replace_date_variables(email_data['message'])

        try:
            if server_type in ["Gmail", "Yahoo"]:
                smtp_server = email_data['server']['smtp_server']
                smtp_port = int(email_data['server']['smtp_port'])
                email_address = email_data['server']['email_address']
                password = email_data['server']['password']

                # Set up the MIME
                msg = MIMEMultipart()
                msg['From'] = email_address
                msg['To'] = email_data['to']
                msg['Subject'] = email_data['subject']
                msg.attach(MIMEText(message_body, 'plain'))  # Use the modified message body

                # Handle Attachments
                for attachment in email_data['attachments']:
                    try:
                        with open(attachment, 'rb') as f:
                            mime_attachment = MIMEText(f.read(), 'base64', 'utf-8')
                            mime_attachment[
                                "Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment)}"'
                            msg.attach(mime_attachment)
                    except Exception as e:
                        print(f"Failed to attach file {attachment}: {str(e)}")

                # Start the SMTP server
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(email_address, password)
                server.sendmail(email_address, email_data['to'], msg.as_string())
                server.quit()

                print(f"Email sent via {server_type}!")  # Log to console instead

            elif server_type == "Outlook":
                self.send_email_via_outlook(email_data)

        except Exception as e:
            print(f"Failed to send email: {str(e)}")  # Log errors to console

    def replace_date_variables(self, message):
        """Replace date-related placeholders in the message with actual values."""
        import re
        from datetime import datetime

        # Current date and time
        current_date = datetime.now()

        # Define a regex pattern to find each date variable
        patterns = {
            r'\{Date\}': current_date.strftime('%m/%d-%Y'),  # Current date
            r'\{DateTime\}': current_date.strftime('%m/%d-%Y %H:%M:%S'),  # Current date and time
            r'\{Day\}': current_date.strftime('%d'),  # Current day (1-31)
            r'\{DayOfWeek\}': current_date.strftime('%A'),  # Current day of the week (Monday, Tuesday, ...)
            r'\{DayOfYear\}': current_date.strftime('%j'),  # Day of the year (1-366)
        }

        # Replace each pattern in the message with its corresponding value
        for pattern, replacement in patterns.items():
            message = re.sub(pattern, replacement, message)

        # Also handle {DateInDays(X)} for subtracting/adding days
        date_in_days_pattern = r'\{DateInDays\((-?\d+)\)\}'
        matches = re.findall(date_in_days_pattern, message)

        for match in matches:
            days = int(match)  # Get the integer value of days
            actual_date = (current_date - timedelta(days)).strftime('%Y-%m-%d')  # Calculate the date
            message = message.replace(f'{{DateInDays({match})}}', actual_date)  # Replace in the message

        return message

    def send_email_via_outlook(self, email_data):
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = email_data['to']
            mail.Subject = email_data['subject']
            mail.Body = email_data['message']

            # Add attachments if any
            for attachment in email_data['attachments']:
                mail.Attachments.Add(attachment)

            mail.Send()
            print("Email sent successfully via Outlook!")  # Log to console instead

        except Exception as e:
            print(f"Failed to send email via Outlook: {str(e)}")  # Log errors to console

    def cancel_email(self):
        self.close()


class EmailAppUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EmailAway")
        self.setWindowIcon(QIcon(r'C:\Users\haft_\Desktop\email Me App\logo.jpg'))
        self.resize(750, 400)
        self.emails = []
        self.timers = {}  # Store active timers for scheduled emails
        self.init_ui()

    def init_login(self):
        with open('data.json', 'r') as f:
            data = json.load(f)
        license_data = data['licenses']
        user_data = data['users']

        login_dialog = LoginDialog(license_data, user_data)
        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            self.init_ui()  # Initialize the rest of the application only on successful login
        else:
            sys.exit()  # Exit the application if the login dialog is closed or canceled

    def init_ui(self):
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        self.email_list = QListWidget(self)
        self.email_list.itemDoubleClicked.connect(self.modify_email)
        layout.addWidget(self.email_list)

        new_email_button = QPushButton("New Email", self)
        new_email_button.clicked.connect(self.create_email_window)
        layout.addWidget(new_email_button)

        delete_email_button = QPushButton("Delete Email", self)
        delete_email_button.clicked.connect(self.delete_selected_email)
        layout.addWidget(delete_email_button)

        self.setCentralWidget(central_widget)
        self.load_emails()
        self.show()

    def create_email_window(self, email_data=None):
        self.email_window = EmailWindow(self, self.add_email_to_list, email_data)
        self.email_window.exec()

    def add_email_to_list(self, email_data):
        # Check if the task already exists, update it if so
        for idx, task in enumerate(self.emails):
            if task['task_name'] == email_data['task_name']:
                self.emails[idx] = email_data
                self.refresh_task_list()
                self.save_emails()  # Save to file after modification
                self.schedule_email(email_data)  # Schedule or re-schedule email
                return

        # If it's a new task, append it and save
        self.emails.append(email_data)
        self.refresh_task_list()
        self.save_emails()  # Save to file after adding
        self.schedule_email(email_data)

    def load_emails(self):
        try:
            with open('emails.json', 'r') as f:
                self.emails = json.load(f)
            self.refresh_task_list()
        except FileNotFoundError:
            print("Emails file not found, starting fresh.")
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON: {e}")

    def refresh_task_list(self):
        """Clear and repopulate the task list with detailed information."""
        self.email_list.clear()

        for email_data in self.emails:
            if isinstance(email_data, dict):
                task_info = f"{email_data.get('task_name', 'Unnamed')} | {email_data['schedule'].get('time', 'No time')} | {email_data['schedule'].get('frequency', 'No frequency')}"
                self.email_list.addItem(task_info)


    def delete_selected_email(self):
        current_row = self.email_list.currentRow()
        if current_row >= 0:
            email_data = self.emails.pop(current_row)
            task_name = email_data['task_name']

            # Cancel the timer if one exists for this task
            if task_name in self.timers:
                self.timers[task_name].cancel()
                del self.timers[task_name]

            self.refresh_task_list()
            self.save_emails()

    def save_emails(self):
        try:
            with open('emails.json', 'w') as file:
                json.dump(self.emails, file, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save emails: {str(e)}")

    def modify_email(self, item):
        task_name = item.text().split(" | ")[0]
        for email_data in self.emails:
            if email_data['task_name'] == task_name:
                self.create_email_window(email_data)
                return

    def schedule_email(self, email_data):
        try:
            # Combine date and time into a single string
            schedule_time_str = f"{email_data['schedule']['start_date']} {email_data['schedule']['time']}"
            schedule_time = datetime.strptime(schedule_time_str, '%m/%d/%Y %I:%M %p')

            # Calculate the delay in seconds
            delay = (schedule_time - datetime.now()).total_seconds()

            if delay < 0:
                QMessageBox.critical(self, "Error", "Scheduled time is in the past!")
                return

            print(f"Scheduling email for: {schedule_time}. Delay: {delay} seconds")

            # Schedule email sending
            timer = threading.Timer(delay, self.send_scheduled_email, args=(email_data,))
            timer.start()

            # Store the timer so we can cancel it if necessary
            self.timers[email_data['task_name']] = timer

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to schedule email: {str(e)}")

    def send_scheduled_email(self, email_data):
        """Send the scheduled email without displaying a success message"""
        print(f"Sending scheduled email to: {email_data['to']} at {datetime.now()}")
        self.send_email(email_data)

    def send_email(self, email_data):
        """Generic method to send email based on the server type"""
        server_type = email_data['server']['server_type']

        try:
            if server_type in ["Gmail", "Yahoo"]:
                smtp_server = email_data['server']['smtp_server']
                smtp_port = int(email_data['server']['smtp_port'])
                email_address = email_data['server']['email_address']
                password = email_data['server']['password']

                # Set up the MIME
                msg = MIMEMultipart()
                msg['From'] = email_address
                msg['To'] = email_data['to']
                msg['Subject'] = email_data['subject']
                body = email_data['message']
                msg.attach(MIMEText(body, 'plain'))

                # Handle Attachments
                for attachment in email_data['attachments']:
                    try:
                        with open(attachment, 'rb') as f:
                            mime_attachment = MIMEText(f.read(), 'base64', 'utf-8')
                            mime_attachment[
                                "Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment)}"'
                            msg.attach(mime_attachment)
                    except Exception as e:
                        print(f"Failed to attach file {attachment}: {str(e)}")

                # Start the SMTP server
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(email_address, password)
                server.sendmail(email_address, email_data['to'], msg.as_string())
                server.quit()

                print(f"Email sent via {server_type}!")

            elif server_type == "Outlook":
                self.send_email_via_outlook(email_data)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send email: {str(e)}")

    def send_email_via_outlook(self, email_data):
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = email_data['to']
            mail.Subject = email_data['subject']
            mail.Body = email_data['message']

            # Add attachments if any
            for attachment in email_data['attachments']:
                mail.Attachments.Add(attachment)

            mail.Send()
            QMessageBox.information(self, "Success", "Email sent successfully via Outlook!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send email via Outlook: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Load user and license data from JSON
    with open('license_data.json', 'r') as f:
        data = json.load(f)
    license_data = data['licenses']
    user_data = data['users']

    # Show the login dialog
    login_dialog = LoginDialog(license_data, user_data)
    if login_dialog.exec() == QDialog.DialogCode.Accepted:
        main_app = EmailAppUI()
        sys.exit(app.exec())  # Start the main event loop here
    else:
        sys.exit()  # Exit the application if the login was unsuccessful
