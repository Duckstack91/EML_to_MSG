import requests
import os
import zipfile
import sys
from tkinter import messagebox
import tkinter as tk
from version import VERSION  # Dies lädt die Version, die du in version.py festgelegt hast

VERSION_URL = "https://raw.githubusercontent.com/Duckstack91/EML_to_MSG/master/src/version.txt"
ZIP_URL = "https://github.com/Duckstack91/EML_to_MSG/archive/refs/heads/master.zip"
LOCAL_VERSION_FILE = "version.txt"


# Funktion zum Abrufen der entfernten Version von GitHub
def get_remote_version():
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()
        return response.text.strip()  # Entferne Leerzeichen und Zeilenumbrüche
    except Exception as e:
        print(f"Fehler beim Abrufen der Version: {e}")
        return None

# Funktion zum Abrufen der lokalen Version
def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return "0.0.0"
def update_version_in_txt():
    try:
        with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(VERSION)  # Schreibe die Version aus version.py in version.txt
        print(f"Version {VERSION} erfolgreich in {LOCAL_VERSION_FILE} geschrieben.")
    except Exception as e:
        print(f"Fehler beim Schreiben der Version in {LOCAL_VERSION_FILE}: {e}")

update_version_in_txt()

# Funktion zur Durchführung des Updates
# Funktion zur Durchführung des Updates
def update_application():
    try:
        response = requests.get(ZIP_URL)
        response.raise_for_status()

        # Herunterladen der ZIP-Datei
        with open("update.zip", "wb") as f:
            f.write(response.content)

        # Entpacken und Installation
        with zipfile.ZipFile("update.zip", "r") as zip_ref:
            zip_ref.extractall(".")
        os.remove("update.zip")

        # Lokale Versionsdatei aktualisieren
        with open(LOCAL_VERSION_FILE, "w") as f:
            f.write(get_remote_version())

        messagebox.showinfo("Update", "Update erfolgreich abgeschlossen. Die neue Version wird gestartet.")

        # Aktuelles Programm neu starten
        python = sys.executable  # Python-Interpreter
        os.execl(python, python, *sys.argv)  # Neustart mit den gleichen Argumenten
    except Exception as e:
        messagebox.showerror("Update-Fehler", f"Fehler beim Update: {e}")
        print(f"Update-Fehler: {e}")


# Funktion zum Konvertieren der Versionsnummer in ein Tupel für den Vergleich
def convert_version_to_tuple(version):
    return tuple(map(int, version.split('.')))

# Funktion zur Überprüfung von Updates
def check_for_updates():
    remote_version = get_remote_version()
    if remote_version:
        local_version = get_local_version()

        print(f"Lokale Version: {local_version}")
        print(f"Entfernte Version: {remote_version}")

        # Vergleiche die Versionsnummern als Tupel
        remote_version_tuple = convert_version_to_tuple(remote_version)
        local_version_tuple = convert_version_to_tuple(local_version)

        # Vergleiche die Versionen
        if remote_version_tuple > local_version_tuple:
            print(f"Eine neue Version {remote_version} ist verfügbar.")
            # Benutzer benachrichtigen und Option zum Updaten anbieten
            response = messagebox.askyesno(
                "Update verfügbar", f"Version {remote_version} ist verfügbar. Möchten Sie das Update durchführen?"
            )
            if response:
                update_application()
        else:
            print("Die aktuelle Version ist die neueste.")
    else:
        print("Fehler beim Abrufen der Version von GitHub.")
        messagebox.showerror("Fehler", "Die Versionsinformation konnte nicht abgerufen werden.")

# Tkinter-GUI zum Starten der Anwendung und zur Anzeige von Nachrichten


