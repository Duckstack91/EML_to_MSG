import requests
import os
import zipfile
import sys
from tkinter import messagebox

VERSION_URL = "https://raw.githubusercontent.com/Duckstack91/EML_to_MSG/main/version.txt"
ZIP_URL = "https://github.com/Duckstack91/EML_to_MSG/archive/refs/heads/main.zip"
LOCAL_VERSION_FILE = "version.txt"

def get_remote_version():
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()
        return response.text.strip()
    except Exception as e:
        print(f"Fehler beim Abrufen der Version: {e}")
        return None

def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return "0.0.0"

def update_application():
    try:
        response = requests.get(ZIP_URL)
        response.raise_for_status()
        with open("update.zip", "wb") as f:
            f.write(response.content)

        with zipfile.ZipFile("update.zip", "r") as zip_ref:
            zip_ref.extractall(".")
        os.remove("update.zip")

        with open(LOCAL_VERSION_FILE, "w") as f:
            f.write(get_remote_version())

        messagebox.showinfo("Update", "Update erfolgreich abgeschlossen. Starten Sie das Programm neu.")
        sys.exit()  # Programm beenden, damit der Nutzer neu starten kann
    except Exception as e:
        messagebox.showerror("Update-Fehler", f"Fehler beim Update: {e}")
        print(f"Update-Fehler: {e}")

def check_for_updates():
    remote_version = get_remote_version()
    local_version = get_local_version()

    if remote_version and remote_version > local_version:
        result = messagebox.askyesno("Update verfügbar", f"Version {remote_version} ist verfügbar. Möchten Sie das Update durchführen?")
        if result:
            update_application()
