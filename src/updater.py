import requests
import os
import sys, re
from tkinter import messagebox
from version import VERSION  # Dies lädt die aktuelle Version

VERSION_URL = "https://raw.githubusercontent.com/Duckstack91/EML_to_MSG/master/src/version.txt"
EXE_URL = "https://github.com/Duckstack91/EML_to_MSG/releases/latest/download/eml_to_msg.exe"
LOCAL_VERSION_FILE = "version.txt"


# Funktion zum Abrufen der entfernten Version von GitHub
# Funktion zum Abrufen der entfernten Version von GitHub
def get_remote_version():
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()  # Überprüft auf HTTP-Fehler
        version = response.text.strip()  # Entfernt Leerzeichen und Zeilenumbrüche

        # Überprüfen, ob die Version dem Format X.Y.Z entspricht
        if re.match(r'^\d+\.\d+\.\d+$', version):  # X.Y.Z Format, z.B. 1.0.1
            return version
        else:
            print(f"Ungültiges Versionsformat empfangen: {version}")
            return None
    except Exception as e:
        print(f"Fehler beim Abrufen der Version: {e}")
        return None


# Funktion zum Abrufen der lokalen Version
def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return "0.0.0"


# Funktion zur Aktualisierung der lokalen Versionsdatei
def update_version_in_txt():
    try:
        with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(VERSION)  # Schreibe die Version aus `version.py` in `version.txt`
        print(f"Version {VERSION} erfolgreich in {LOCAL_VERSION_FILE} geschrieben.")
    except Exception as e:
        print(f"Fehler beim Schreiben der Version in {LOCAL_VERSION_FILE}: {e}")


# Funktion zur Durchführung des Updates
def update_application():
    try:
        # Herunterladen der neuen .exe
        response = requests.get(EXE_URL, stream=True)
        response.raise_for_status()

        # Temporäre Datei für die neue Version
        new_exe = "update_new.exe"
        with open(new_exe, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        # Aktuelles ausführbares Programm ersetzen
        current_exe = sys.executable
        backup_exe = current_exe + ".old"

        os.rename(current_exe, backup_exe)  # Aktuelles Programm sichern
        os.rename(new_exe, current_exe)    # Neue Version umbenennen

        messagebox.showinfo("Update", "Update erfolgreich abgeschlossen. Bitte starten Sie die Anwendung neu.")
        sys.exit()  # Beenden der aktuellen Instanz
    except Exception as e:
        messagebox.showerror("Update-Fehler", f"Fehler beim Update: {e}")
        print(f"Update-Fehler: {e}")


# Funktion zum Konvertieren der Versionsnummer in ein Tupel für den Vergleich
def convert_version_to_tuple(version):
    return tuple(map(int, version.split('.')))


# Funktion zur Überprüfung von Updates
def check_for_updates():
    remote_version = get_remote_version()

    if remote_version is None:
        print("Update-Überprüfung übersprungen (Netzwerkfehler oder ungültige Version).")
        return  # Überspringe die Versionsprüfung

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


# Aktualisiere die lokale Version beim Start
update_version_in_txt()
