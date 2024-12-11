import json
import uuid
from datetime import datetime, timedelta
import hashlib
import os
import tkinter as tk
from tkinter import ttk, messagebox

# File path to the JSON file where licenses and users are stored
LICENSE_FILE = "license_data.json"

def load_data():
    """Load data from the JSON file."""
    if not os.path.exists(LICENSE_FILE):
        messagebox.showerror("Error", f"File '{LICENSE_FILE}' does not exist. Please run the main application to create the file.")
        return None
    with open(LICENSE_FILE, 'r') as file:
        return json.load(file)

def save_data(data):
    """Save data to the JSON file."""
    with open(LICENSE_FILE, 'w') as file:
        json.dump(data, file, indent=4)

def hash_password(password):
    """Hash the password using SHA256."""
    return hashlib.sha256(password.encode()).hexdigest()

class LicenseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("License Management System")
        self.geometry("400x300")
        self.create_widgets()

    def create_widgets(self):
        # Username Entry
        # Username Entry
        ttk.Label(self, text="Username:").grid(row=0, column=0, padx=10, pady=10)
        self.username_entry = ttk.Entry(self)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)

        # Password Entry
        ttk.Label(self, text="Password:").grid(row=1, column=0, padx=10, pady=10)
        self.password_entry = ttk.Entry(self, show="*")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        # Submit Button for License Creation
        ttk.Button(self, text="Create License", command=self.create_license).grid(row=2, column=0, columnspan=2,
                                                                                  pady=10)


    def create_license(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        if not username or not password:
            messagebox.showerror("Error", "Username and password cannot be empty")
            return

        data = load_data()
        if data is None:
            return

        # Check if the user already exists
        if any(user["username"] == username for user in data["users"]):
            messagebox.showerror("Error", f"User '{username}' already exists.")
            return

        # Create a new license
        license_key = str(uuid.uuid4()).upper()
        expiration_date = (datetime.now() + timedelta(days=365)).isoformat()  # One year from now

        new_license = {
            "license_key": license_key,
            "expiration_date": expiration_date,
            "max_users": 5,
            "active": True,
            "subscription_paid": True
        }

        password_hash = hash_password(password)
        new_user = {
            "username": username,
            "password_hash": password_hash,
            "license_key": license_key,
            "last_login": None,
            "start_date": datetime.now().isoformat(),  # Marking start date as current date
            "role": "active"  # Initial role set as active
        }

        # Add the license and user to the JSON data
        data["licenses"].append(new_license)
        data["users"].append(new_user)

        # Save the updated data to the JSON file
        save_data(data)
        messagebox.showinfo("Success", f"Successfully created license for user '{username}' with license key '{license_key}'.")



if __name__ == "__main__":
    app = LicenseApp()
    app.mainloop()
