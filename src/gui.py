import tkinter as tk
from tkinter import filedialog, messagebox
import os
from configparser import ConfigParser
import traceback
from eml_to_msg import process_directory
from updater import check_for_updates
import sys

CONFIG_FILE = 'config.ini'
current_dir = os.path.dirname(os.path.realpath(__file__))

# version.py
VERSION = "1.0.1"
##Version Ändern für Release
##git tag -a v1.0.2 -m "Release version 1.0.2"
##git push origin v1.0.2


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

    def convert(self):
        self.eml_directory = self.eml_dir_entry.get()
        self.output_directory = self.config['directories']['output_directory']

        if not self.eml_directory or not self.output_directory:
            messagebox.showerror("Fehler", "Bitte wählen Sie sowohl ein EML- als auch ein Zielverzeichnis aus.")
            return

        if not os.path.isdir(self.eml_directory) or not os.path.isdir(self.output_directory):
            messagebox.showerror("Fehler", "Ein ausgewähltes Verzeichnis existiert nicht.")
            return

        try:
            process_directory(self.eml_directory, self.output_directory)
            messagebox.showinfo("Erfolg", "Die Konvertierung wurde erfolgreich abgeschlossen.")
            self.save_config()
        except Exception as e:
            error_message = f"Ein Fehler ist aufgetreten: {str(e)}"
            print(f"{error_message}\n\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", error_message)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            self.config.read(CONFIG_FILE)
        else:
            self.config['directories'] = {'eml_directory': '', 'output_directory': ''}

    def save_config(self):
        self.config['directories']['eml_directory'] = self.eml_directory
        with open(CONFIG_FILE, 'w') as configfile:
            self.config.write(configfile)

if __name__ == "__main__":
    check_for_updates()  # Updates vor GUI-Start prüfen
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
