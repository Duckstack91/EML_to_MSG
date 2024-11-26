import tkinter as tk
from tkinter import filedialog, messagebox
import os
from configparser import ConfigParser
import traceback
from eml_to_msg import process_directory
from updater import check_for_updates
import sys

CONFIG_FILE = os.path.join(os.getenv('APPDATA'), 'EML_to_MSG', 'config.ini')
current_dir = os.path.dirname(os.path.realpath(__file__))

# version.py
VERSION = "1.0.1"
##Version Ändern für Release
##git tag -a v1.0.3 -m "Release version 1.0.3"
##git push origin v1.0.3
##pyinstaller --onefile --windowed --version-file=version.txt --name Eml_to_Msg.exe gui.py
def ensure_config_exists():
    """Überprüft, ob die config.ini existiert, und erstellt sie mit Standardwerten, falls nicht."""
    config_dir = os.path.dirname(CONFIG_FILE)
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)

    if not os.path.exists(CONFIG_FILE):
        print(f"{CONFIG_FILE} nicht gefunden. Erstelle Standardkonfiguration...")
        config = ConfigParser()
        config['directories'] = {
            'eml_directory': '',
            'output_directory': ''
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        print(f"{CONFIG_FILE} wurde erfolgreich erstellt.")
    else:
        print(f"{CONFIG_FILE} existiert bereits.")

    # Repariere Konfigurationsdatei, falls Sektionen/Schlüssel fehlen
    config = ConfigParser()
    config.read(CONFIG_FILE)

    if 'directories' not in config:
        config['directories'] = {}
    if 'eml_directory' not in config['directories']:
        config['directories']['eml_directory'] = ''
    if 'output_directory' not in config['directories']:
        config['directories']['output_directory'] = ''

    # Änderungen speichern
    with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

class ToolTip:
    """Toofltip class to display tooltips for wjjidgets."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.tooltip_id = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        self.tooltip_id = self.widget.after(5000, self._create_tooltip)

    def _create_tooltip(self):
        if self.tooltip_window is None:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + 20
            self.tooltip_window = tk.Toplevel(self.widget)
            self.tooltip_window.wm_overrideredirect(True)
            self.tooltip_window.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip_window, text=self.text, background="lightyellow", borderwidth=1, relief="solid")
            label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        if self.tooltip_id:
            self.widget.after_cancel(self.tooltip_id)
            self.tooltip_id = None

class ConverterApp:
    def save_config(self):
        """Speichert die aktuellen Konfigurationswerte in der config.ini."""
        try:
            self.config['directories']['eml_directory'] = self.eml_directory
            self.config['directories']['output_directory'] = self.output_directory

            # Pfad sicherstellen und Datei speichern
            config_dir = os.path.dirname(CONFIG_FILE)
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)

            with open(CONFIG_FILE, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
            print(f"Konfigurationsdatei wurde erfolgreich aktualisiert: {CONFIG_FILE}")
        except Exception as e:
            print(f"Fehler beim Speichern der Konfiguration: {e}")
            traceback.print_exc()

    def __init__(self, root):
        self.root = root
        self.root.title("EML to MSG Converter")
        current_directory = os.path.dirname(os.path.abspath(__file__))

       # if getattr(sys, 'frozen', False):
           # icon_path = os.path.join(sys._MEIPASS, 'icons', 'Screenshot_1.png')
        #else:
         #   icon_path = os.path.join(current_directory, 'icons', 'Screenshot_1.png')

        #icon = tk.PhotoImage(file=icon_path)
       # self.root.iconphoto(False, icon)
        #print(f"{icon_path}")
        self.config = ConfigParser()
        self.load_config()

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="EML Verzeichnis auswählen:").pack(pady=5)
        self.eml_dir_entry = tk.Entry(self.root, width=50)
        self.eml_dir_entry.pack(padx=5)

        self.browse_eml_button = tk.Button(self.root, text="Durchsuchen", command=self.browse_eml_directory)
        self.browse_eml_button.pack(pady=5)

        ToolTip(self.browse_eml_button, "Wählen Sie das Verzeichnis mit EML-Dateien aus.")

        self.browse_eml_button.bind("<Enter>", self.on_enter_browse)
        self.browse_eml_button.bind("<Leave>", self.on_leave_browse)

        self.convert_button = tk.Button(self.root, text="Konvertieren", command=self.convert)
        self.convert_button.pack(pady=10)

        ToolTip(self.convert_button, "Klick hier, um die Konvertierung von EML zu MSG zu starten.")

        self.convert_button.bind("<Enter>", self.on_enter_convert)
        self.convert_button.bind("<Leave>", self.on_leave_convert)

        self.eml_dir_entry.insert(0, self.config['directories']['eml_directory'])

        root.resizable(False, False)

    def on_enter_browse(self, event):
        self.browse_eml_button.config(bg='lightgrey')

    def on_leave_browse(self, event):
        self.browse_eml_button.config(bg='SystemButtonFace')

    def on_enter_convert(self, event):
        self.convert_button.config(bg='lightgrey')

    def on_leave_convert(self, event):
        self.convert_button.config(bg='SystemButtonFace')

    def browse_eml_directory(self):
        self.eml_directory = filedialog.askdirectory()
        if self.eml_directory:
            self.eml_dir_entry.delete(0, tk.END)
            self.eml_dir_entry.insert(0, self.eml_directory)

            # Automatisch den Output-Ordner festlegen
            self.output_directory = os.path.join(self.eml_directory, "_EmlKonvertiert")
            if not os.path.exists(self.output_directory):
                os.makedirs(self.output_directory)
            #print(f"Output-Ordner gesetzt auf: {self.output_directory}")

    def convert(self):
        self.eml_directory = self.eml_dir_entry.get()
        if not self.eml_directory:
            messagebox.showerror("Fehler", "Bitte wählen Sie ein EML-Verzeichnis aus.")
            return

        if not os.path.isdir(self.eml_directory):
            messagebox.showerror("Fehler", "Das ausgewählte Quellverzeichnis existiert nicht.")
            return

        try:
            # Konvertierung starten
            process_directory(self.eml_directory, self.output_directory)
            messagebox.showinfo("Erfolg", f"Die  Konvertierung wurde erfolgreich abgeschlossen.\n"
                                          f"Dateien gespeichert in: {self.output_directory}")
            self.save_config()  # Konfiguration speichern
        except Exception as e:
            error_message = f"Ein Fehler ist aufgetreten: {str(e)}"
            print(f"{error_message}\n\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", error_message)


    def load_config(self):
        """Lädt die Konfiguration und stellt sicher, dass alle erforderlichen Schlüssel vorhanden sind."""
        if os.path.exists(CONFIG_FILE):
            self.config.read(CONFIG_FILE)
        else:
            self.config['directories'] = {'eml_directory': '', 'output_directory': ''}
            return

        # Sicherstellen, dass die Sektion 'directories' existiert
        if 'directories' not in self.config:
            self.config['directories'] = {}

        # Standardwerte für fehlende Schlüssel setzen
        if 'eml_directory' not in self.config['directories']:
            self.config['directories']['eml_directory'] = ''
        if 'output_directory' not in self.config['directories']:
            self.config['directories']['output_directory'] = ''


if __name__ == "__main__":
    check_for_updates()  # Updates vor GUI-Start prüfen
    ensure_config_exists()
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
