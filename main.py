import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import threading
import json
import imaplib
import email
import os
import sv_ttk

# Define file paths
SETTINGS_FILE = "settings.json"

def save_settings(settings):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=4)

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as f:
            return json.load(f)
    else:
        return {}

# Load settings
settings = load_settings()

# If settings are empty, save the initial settings
if not settings:
    settings = {}  # Initialize empty settings
    save_settings(settings)

# GUI-related code
class CustomDialog(tk.Toplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        label = tk.Label(self, text=message)
        label.pack(padx=20, pady=10)
        ok_button = tk.Button(self, text="OK", command=self.destroy)
        ok_button.pack(pady=5)

class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, settings):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("800x600")
        self["padx"] = 10
        self["pady"] = 10

        self.settings = settings
        self.folder_destinations = {}

        self.imap_server_label = tk.Label(self, text="IMAP strežnik:")
        self.imap_server_label.grid(row=0, column=0, sticky="w")
        self.imap_server_entry = tk.Entry(self)
        self.imap_server_entry.grid(row=0, column=1, sticky="ew")

        self.email_label = tk.Label(self, text="Email naslov:")
        self.email_label.grid(row=1, column=0, sticky="w")
        self.email_entry = tk.Entry(self)
        self.email_entry.grid(row=1, column=1, sticky="ew")

        self.password_label = tk.Label(self, text="Geslo:")
        self.password_label.grid(row=2, column=0, sticky="w")
        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.grid(row=2, column=1, sticky="ew")

        self.connect_button = tk.Button(self, text="Poveži", command=self.connect_to_imap)
        self.connect_button.grid(row=3, column=0, columnspan=2, pady=10)

        label1 = tk.Label(self, text="IMAP mapa", anchor="center")
        label1.grid(row=4, column=0, sticky="ew")

        label2 = tk.Label(self, text="Ciljna mapa", anchor="center")
        label2.grid(row=4, column=1, sticky="ew")

        label3 = tk.Label(self, text="Izbrana mapa", anchor="center")
        label3.grid(row=4, column=2, sticky="ew")

        separator = ttk.Separator(self, orient="horizontal", style="black.Horizontal.TSeparator")
        separator.grid(row=5, columnspan=3, sticky="ew", pady=5)

        self.canvas = tk.Canvas(self, borderwidth=0)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.grid(row=6, column=0, columnspan=3, sticky="nsew")
        self.scrollbar.grid(row=6, column=3, sticky="ns")

        self.grid_rowconfigure(6, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.save_button = tk.Button(self, text="Shrani", command=self.save_settings)
        self.save_button.grid(row=7, columnspan=2)

        # Load settings from the stored settings
        self.imap_server_entry.insert(0, settings.get("imap_server", ""))
        self.email_entry.insert(0, settings.get("email_address", ""))
        self.password_entry.insert(0, settings.get("password", ""))

        # Populate the folder destinations
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
                label = tk.Label(self.scrollable_frame, text=folder_name)
                label.grid(row=idx, column=0, sticky="w")

                entry = tk.Entry(self.scrollable_frame)
                entry.grid(row=idx, column=1)

                browse_button = tk.Button(
                    self.scrollable_frame,
                    text="Prebrskaj",
                    command=lambda folder=folder_name, entry=entry: self.add_destination(folder, entry),
                )
                browse_button.grid(row=idx, column=2)

                checkbox_var = tk.BooleanVar()
                checkbox = tk.Checkbutton(self.scrollable_frame, variable=checkbox_var)
                checkbox.grid(row=idx, column=3, sticky="ew")

                self.folder_destinations[folder_name] = {
                    "entry": entry,
                    "selected": checkbox_var,
                }

            CustomDialog(
                self,
                "Uspešno povezano",
                "Povezava s serverjem je uspešno vzpostavljena!",
            )
            mail.logout()
        except Exception as e:
            CustomDialog(
                self,
                "Napaka pri povezavi",
                f"Prišlo je do napake pri povezovanju na strežnik: {str(e)}",
            )

    def add_destination(self, folder, entry):
        destination = filedialog.askdirectory()
        if destination:
            entry.delete(0, tk.END)
            entry.insert(0, destination)
            self.folder_destinations[folder]["entry"] = entry
            checkbox_var = self.folder_destinations[folder]["selected"]
            self.folder_destinations[folder]["selected"] = checkbox_var

    def populate_folder_destinations(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        folder_destinations = self.settings.get("folder_destinations", {})
        for idx, (folder, destination_data) in enumerate(folder_destinations.items()):
            label = tk.Label(self.scrollable_frame, text=folder, anchor="center")
            label.grid(row=idx, column=0, sticky="w")

            entry = tk.Entry(self.scrollable_frame)
            entry.insert(0, destination_data.get("destination", ""))
            entry.grid(row=idx, column=1)

            browse_button = tk.Button(
                self.scrollable_frame,
                text="Prebrskaj",
                command=lambda f=folder, e=entry: self.add_destination(f, e),
            )
            browse_button.grid(row=idx, column=2)

            checkbox_var = tk.BooleanVar(value=destination_data.get("selected", False))
            checkbox = tk.Checkbutton(self.scrollable_frame, variable=checkbox_var)
            checkbox.grid(row=idx, column=3)

            self.folder_destinations[folder] = {
                "entry": entry,
                "selected": checkbox_var,
            }

    def save_settings(self):
        settings = {
            "imap_server": self.imap_server_entry.get(),
            "email_address": self.email_entry.get(),
            "password": self.password_entry.get(),
            "folder_destinations": {},
        }
        for folder, data in self.folder_destinations.items():
            selected = data["selected"].get()
            if selected is None:
                selected = False
            settings["folder_destinations"][folder] = {
                "destination": data["entry"].get(),
                "selected": selected,
            }
        try:
            save_settings(settings)
            CustomDialog(root, "Shranjeno", "Nastavitve so bile uspešno shranjene!")
        except Exception as e:
            CustomDialog(root, "Napaka", f"Prišlo je do napake: {str(e)}")

def fetch_emails():
    global fetching_should_continue
    fetching_should_continue = True
    settings = load_settings()

    imap_server = settings.get("imap_server")
    email_address = settings.get("email_address")
    password = settings.get("password")
    folder_destinations = settings.get("folder_destinations", {})

    file_types_entry_value = file_types_entry.get()
    file_types = (
        [x.strip() for x in file_types_entry_value.split(",")]
        if file_types_entry_value
        else []
    )

    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, password)

        for folder, folder_data in folder_destinations.items():
            if folder_data.get("selected", False):
                mail.select(folder)
                _, data = mail.search(None, "ALL")

                if data:
                    num_messages = len(data[0].split())
                    progress_step = 100 / num_messages
                    current_progress = 0
                    for num in data[0].split():
                        if not fetching_should_continue:
                            break
                        _, data = mail.fetch(num, "(RFC822)")

                        if data:
                            raw_email = data[0][1]
                            email_message = email.message_from_bytes(raw_email)

                            for part in email_message.walk():
                                if part.get_content_maintype() == "multipart":
                                    continue
                                if part.get("Content-Disposition") is None:
                                    continue
                                filename = part.get_filename()

                                if filename:
                                    file_extension = os.path.splitext(filename)[1][1:]
                                    if not file_types or any(
                                        filename.lower().endswith(file_type.lower())
                                        for file_type in file_types
                                    ):
                                        destination_folder = folder_data.get("destination")
                                        if destination_folder:
                                            if not os.path.exists(destination_folder):
                                                os.makedirs(destination_folder)
                                            filepath = os.path.join(
                                                destination_folder, filename
                                            )
                                            with open(filepath, "wb") as f:
                                                f.write(part.get_payload(decode=True))
                        current_progress += progress_step
                        progress_bar["value"] = current_progress
                        root.update_idletasks()

        CustomDialog(
            root, "Končano", "Priponke so bile uspešno prenešene!"
        )
        start_button.grid()
        stop_button.grid_remove()
    except Exception as e:
        CustomDialog(root, "Napaka", f"Prišlo je do napake pri prenosu priponk: {str(e)}")
    finally:
        if mail:
            mail.close()
            mail.logout()
        fetching_should_continue = False
        progress_bar["value"] = 0

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

def save_settings_to_file(updated_settings):
    try:
        # Load the current settings
        current_settings = load_settings()

        # Update only the specific setting passed in updated_settings
        current_settings.update(updated_settings)

        # Save the updated settings
        save_settings(current_settings)
        CustomDialog(root, "Shranjeno", "Nastavitve so bile uspešno shranjene!")
    except Exception as e:
        CustomDialog(root, "Napaka", f"Prišlo je do napake pri shranjenvanju nastavitev: {str(e)}")

def update_file_types():
    new_file_types = file_types_entry.get()
    # Only update the file_types setting, preserving the rest
    save_settings_to_file({"file_types": new_file_types.split(",") if new_file_types else []})

def open_settings_window():
    settings = load_settings()  # Reload settings to ensure the latest values are shown
    settings_window = SettingsWindow(root, settings)

# Initialize the main application window
root = tk.Tk()
root.title("Prenos email priponk")

fetching_should_continue = True

fetch_button_text = tk.StringVar()
fetch_button_text.set("Prenesi")

style = ttk.Style()
style.theme_use("default")
style.configure("Custom.Horizontal.TProgressbar", thickness=50)

root["padx"] = 10
root["pady"] = 5

# Place the settings button on the grid
settings_button = tk.Button(root, text="Nastavitve", command=open_settings_window)
settings_button.grid(row=0, column=1, sticky="ne", pady=20)

# Initialize file types entry field
file_types_label = tk.Label(root, text="Tip datoteke za prenos (ločeno z vejico)")
file_types_label.grid(row=1, column=0, sticky="w")
file_types_entry = tk.Entry(root)
file_types_entry.insert(0, ",".join(settings.get("file_types", [])))
file_types_entry.grid(row=1, column=1)

update_button = tk.Button(root, text="Shrani", command=update_file_types)
update_button.grid(row=1, column=2)

file_types_description = tk.Label(
    root,
    text="Pustite prazno, če želite prenesti vse datoteke v priponkah e-pošte.",
    font=("Helvetica", 13, "italic"),
)
file_types_description.grid(row=2, column=0, columnspan=2, sticky="w")

start_button = tk.Button(root, text="Prenesi priponke", command=start_fetching)
start_button.grid(row=3, column=0, columnspan=2, pady=20)

stop_button = tk.Button(root, text="Prekliči prenos", command=stop_fetching)
stop_button.grid(row=3, column=0, columnspan=2, pady=10)
stop_button.grid_remove()

progress_bar = tk.ttk.Progressbar(
    root,
    orient="horizontal",
    length=500,
    mode="determinate",
    style="Custom.Horizontal.TProgressbar",
)
progress_bar.grid(row=4, columnspan=2, pady=30)

status_label = tk.Label(root, text="", anchor="center")
status_label.grid(row=5, columnspan=2, pady=10)

APP_VERSION = "v1.1"
APP_AUTHOR = "BlazL"

app_info_label = tk.Label(root, text=f"Verzija: {APP_VERSION} | Avtor: {APP_AUTHOR}")
app_info_label.grid(row=6, columnspan=2, pady=10)

sv_ttk.set_theme("light")

# Start the Tkinter main loop
root.mainloop()
