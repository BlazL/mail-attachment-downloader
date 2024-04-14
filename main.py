import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from cryptography.fernet import Fernet
import sv_ttk
import json
import imaplib
import email
import os
import threading
import gettext

KEY_FILE = 'key.key'

def generate_key():
    if not os.path.exists(KEY_FILE):
        key = Fernet.generate_key()
        with open(KEY_FILE, 'wb') as key_file:
            key_file.write(key)

def load_key():
    with open(KEY_FILE, 'rb') as key_file:
        return key_file.read()

generate_key()
key = load_key()
cipher_suite = Fernet(key)

def encrypt_settings(settings):
    encrypted_settings = cipher_suite.encrypt(json.dumps(settings).encode())
    with open('settings.json', 'wb') as f:
        f.write(encrypted_settings)

def decrypt_settings():
    try:
        with open('settings.json', 'rb') as f:
            encrypted_settings = f.read()
        decrypted_settings = cipher_suite.decrypt(encrypted_settings).decode()
        return json.loads(decrypted_settings)
    except FileNotFoundError:
        return {}

if not os.path.exists('settings.json'):
    settings = {}  # Initialize empty settings
    encrypt_settings(settings)

settings = decrypt_settings()

class CustomDialog(tk.Toplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)

        label = tk.Label(self, text=message)
        label.pack(padx=20, pady=10)

        ok_button = tk.Button(self, text=_("OK"), command=self.destroy)
        ok_button.pack(pady=5)

class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, settings):
        super().__init__(parent)
        self.title(_("Settings"))

        self['padx'] = 10
        self['pady'] = 5

        self.settings = settings

        self.folder_destinations = {}

        self.imap_server_label = tk.Label(self, text=_("IMAP Server:"))
        self.imap_server_label.grid(row=0, column=0, sticky="w")
        self.imap_server_entry = tk.Entry(self)
        self.imap_server_entry.grid(row=0, column=1)

        self.email_label = tk.Label(self, text=_("Email Address:"))
        self.email_label.grid(row=1, column=0, sticky="w")
        self.email_entry = tk.Entry(self)
        self.email_entry.grid(row=1, column=1)

        self.password_label = tk.Label(self, text=_("Password:"))
        self.password_label.grid(row=2, column=0, sticky="w")
        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.grid(row=2, column=1)

        self.connect_button = tk.Button(self, text=_("Connect"), command=self.connect_to_imap)
        self.connect_button.grid(row=3, columnspan=2, pady=10)

        label1 = tk.Label(self, text=_("IMAP Folder"), anchor="center")
        label1.grid(row=4, column=0, sticky="w")

        label2 = tk.Label(self, text=_("Destination Folder"), anchor="center")
        label2.grid(row=4, column=1, sticky="w")
        label2 = tk.Label(self, text=_("Selected"), anchor="center")
        label2.grid(row=4, column=2, sticky="w")

        separator = ttk.Separator(self, orient='horizontal', style="black.Horizontal.TSeparator")
        separator.grid(row=5, columnspan=3, sticky="ew", pady=5)

        self.settings_frame = tk.Frame(self)
        self.settings_frame.grid(row=6, columnspan=2, sticky="nsew")

        self.save_button = tk.Button(self, text=_("Save"), command=self.save_settings)
        self.save_button.grid(row=7, columnspan=2)

        self.imap_server_entry.insert(0, settings.get("imap_server", ""))
        self.email_entry.insert(0, settings.get("email_address", ""))
        self.password_entry.insert(0, settings.get("password", ""))

        self.populate_folder_destinations()

    def connect_to_imap(self):
        imap_server = self.imap_server_entry.get()
        email_address = self.email_entry.get()
        password = self.password_entry.get()

        try:
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_address, password)
            _, folders = mail.list()
            for idx, folder in enumerate(folders):
                folder_name = folder.decode().split(' "/" ')[1]
                label = tk.Label(self.settings_frame, text=folder_name)
                label.grid(row=idx, column=0, sticky="w")

                entry = tk.Entry(self.settings_frame)
                entry.grid(row=idx, column=1)

                browse_button = tk.Button(self.settings_frame, text=_("Browse"), command=lambda folder=folder_name, entry=entry: self.add_destination(folder, entry))
                browse_button.grid(row=idx, column=2)

                checkbox_var = tk.BooleanVar()  # Create BooleanVar for checkbox
                checkbox = tk.Checkbutton(self.settings_frame, variable=checkbox_var)
                checkbox.grid(row=idx, column=3)

                self.folder_destinations[folder_name] = {"entry": entry, "selected": checkbox_var}

            CustomDialog(self, _("Successful Connection"), _("Connection to the server was successful!"))
            mail.logout()
        except Exception as e:
            CustomDialog(self, _("Connection Error"), _("Error connecting to the server: {str(e)}"))

    def add_destination(self, folder, entry):
        destination = filedialog.askdirectory()
        if destination:
            entry.delete(0, tk.END)
            entry.insert(0, destination)
            self.folder_destinations[folder]["entry"] = entry
            checkbox_var = self.folder_destinations[folder]["selected"]
            self.folder_destinations[folder]["selected"] = checkbox_var

    def populate_folder_destinations(self):
        for widget in self.settings_frame.winfo_children():
            widget.destroy()

        folder_destinations = self.settings.get("folder_destinations", {})
        for idx, (folder, destination_data) in enumerate(folder_destinations.items()):
            label = tk.Label(self.settings_frame, text=folder)
            label.grid(row=idx, column=0, sticky="w")

            entry = tk.Entry(self.settings_frame)
            entry.insert(0, destination_data.get("destination", ""))  # Retrieve the destination path
            entry.grid(row=idx, column=1)

            browse_button = tk.Button(self.settings_frame, text=_("Browse"), command=lambda f=folder, e=entry: self.add_destination(f, e))
            browse_button.grid(row=idx, column=2)

            checkbox_var = tk.BooleanVar(value=destination_data.get("selected", False))  # Retrieve the selected value
            checkbox = tk.Checkbutton(self.settings_frame, variable=checkbox_var)
            checkbox.grid(row=idx, column=3)

            self.folder_destinations[folder] = {"entry": entry, "selected": checkbox_var}

    def save_settings(self):
        settings = {
            "imap_server": self.imap_server_entry.get(),
            "email_address": self.email_entry.get(),
            "password": self.password_entry.get(),
            "folder_destinations": {}
        }
        for folder, data in self.folder_destinations.items():
            selected = data["selected"].get()
            if selected is None:
                selected = False
            settings["folder_destinations"][folder] = {
                "destination": data["entry"].get(),
                "selected": selected
            }
        try:
            encrypted_settings = cipher_suite.encrypt(json.dumps(settings).encode())
            with open('settings.json', 'wb') as f:
                f.write(encrypted_settings)
            CustomDialog(root, _("Saved"), _("Settings have been successfully saved!"))
        except Exception as e:
            CustomDialog(root, _("Error"), _("Error saving settings: {str(e)}"))

def fetch_emails():
    global fetching_should_continue
    fetching_should_continue = True
    try:
        with open('settings.json', 'rb') as f:
            encrypted_settings = f.read()
            decrypted_settings = cipher_suite.decrypt(encrypted_settings).decode()
            settings = json.loads(decrypted_settings)
    except FileNotFoundError:
        CustomDialog(root, _("Error"), _("Cannot find settings.json file"))
        return

    imap_server = settings.get("imap_server")
    email_address = settings.get("email_address")
    password = settings.get("password")

    folder_destinations = settings.get("folder_destinations", {})

    file_types_entry_value = file_types_entry.get()
    file_types = [x.strip() for x in file_types_entry_value.split(",")] if file_types_entry_value else []

    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, password)

        for folder, folder_data in folder_destinations.items():
            if folder_data.get("selected", False):
                mail.select(folder)
                _, data = mail.search(None, 'ALL')

                if data:
                    num_messages = len(data[0].split())
                    progress_step = 100 / num_messages
                    current_progress = 0
                    for num in data[0].split():
                        if not fetching_should_continue:
                            break
                        _, data = mail.fetch(num, '(RFC822)')

                        if data:
                            raw_email = data[0][1]
                            email_message = email.message_from_bytes(raw_email)

                            for part in email_message.walk():
                                if part.get_content_maintype() == 'multipart':
                                    continue
                                if part.get('Content-Disposition') is None:
                                    continue
                                filename = part.get_filename()

                                if filename:
                                    file_extension = os.path.splitext(filename)[1][1:]
                                    if not file_types or any(filename.lower().endswith(file_type.lower()) for file_type in file_types):
                                       destination_folder = folder_data.get("destination")
                                       if destination_folder:
                                           if not os.path.exists(destination_folder):
                                               os.makedirs(destination_folder)
                                           filepath = os.path.join(destination_folder, filename)
                                           with open(filepath, 'wb') as f:
                                               f.write(part.get_payload(decode=True))
                        current_progress += progress_step
                        progress_bar['value'] = current_progress
                        root.update_idletasks()

        CustomDialog(root, _("Finished"), _("Attachment download completed successfully!"))
        start_button.grid()
        stop_button.grid_remove()
    except Exception as e:
         CustomDialog(root, _("Error"),  _("Error downloading attachments: {str(e)}"))
    finally:
        if mail:
            mail.close()
            mail.logout()
        fetching_should_continue = False
        progress_bar['value'] = 0

def start_fetching():
    threading.Thread(target=fetch_emails).start()
    toggle_buttons()

def stop_fetching():
    global fetching_should_continue
    fetching_should_continue = False
    toggle_buttons()

def toggle_buttons():
    if fetching_should_continue:
        start_button.grid_remove()
        stop_button.grid()
    else:
        stop_button.grid_remove()
        start_button.grid()

def save_settings_to_file(settings):
    try:
        encrypted_settings = cipher_suite.encrypt(json.dumps(settings).encode())
        with open('settings.json', 'wb') as f:
            f.write(encrypted_settings)
        CustomDialog(root, _("Saved"), _("Settings have been successfully saved!"))
    except Exception as e:
        CustomDialog(root, _("Error"), _("Error saving settings: {str(e)}"))

def update_file_types():
    new_file_types = file_types_entry.get()
    settings["file_types"] = new_file_types.split(",") if new_file_types else []
    save_settings_to_file(settings)

def open_settings_window():
    try:
        with open('settings.json', 'rb') as f:
            encrypted_settings = f.read()
        decrypted_settings = cipher_suite.decrypt(encrypted_settings).decode()
        settings = json.loads(decrypted_settings)
    except FileNotFoundError:
        settings = {}
    settings_window = SettingsWindow(root, settings)

root = tk.Tk()
root.title(_("Email Attachment Downloader"))

fetching_should_continue = True

fetch_button_text = tk.StringVar()
fetch_button_text.set(_("Fetch Emails"))

style = ttk.Style()
style.theme_use('default')
style.configure("Custom.Horizontal.TProgressbar", thickness=50)

root['padx'] = 10
root['pady'] = 5

settings_button = tk.Button(root, text=_("Settings"), command=open_settings_window)
settings_button.grid(row=0, column=1, sticky="ne", pady=20)

file_types_label = tk.Label(root, text=_("File type to download (comma-separated):"))
file_types_label.grid(row=1, column=0, sticky="w")
file_types_entry = tk.Entry(root)
file_types_entry.insert(0, ",".join(settings.get("file_types", [])))
file_types_entry.grid(row=1, column=1)

update_button = tk.Button(root, text=_("Save"), command=update_file_types)
update_button.grid(row=1, column=2)

file_types_description = tk.Label(root, text=_("Leave empty to download all files in email attachments."), font=("Helvetica", 13, "italic"))
file_types_description.grid(row=2, column=0, columnspan=2, sticky="w")

start_button = tk.Button(root, text=_("Download Attachments"), command=start_fetching)
start_button.grid(row=3, column=0, columnspan=2, pady=20)

stop_button = tk.Button(root, text=_("Stop Download"), command=stop_fetching)
stop_button.grid(row=3, column=0, columnspan=2, pady=10)
stop_button.grid_remove()

progress_bar = tk.ttk.Progressbar(root, orient='horizontal', length=500, mode='determinate', style="Custom.Horizontal.TProgressbar")
progress_bar.grid(row=4, columnspan=2, pady=30)

status_label = tk.Label(root, text="", anchor="center")
status_label.grid(row=5, columnspan=2, pady=10)

app_version = "v1.0"
app_author = "Blaž Lapanja (blaz.lapanja@gmail.com)"

app_info_label = tk.Label(root, text=f"Version: {app_version} | Author: {app_author}")
app_info_label.grid(row=6, columnspan=2, pady=10)

sv_ttk.set_theme("light")

root.mainloop()