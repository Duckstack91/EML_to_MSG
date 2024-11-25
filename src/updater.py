def update_application():
    try:
        # Laden Sie die neue .exe-Datei herunter
        exe_url = "https://github.com/Duckstack91/EML_to_MSG/releases/latest/download/eml_to_msg.exe"
        response = requests.get(exe_url, stream=True)
        response.raise_for_status()

        # Speichern der heruntergeladenen Datei unter einem temporären Namen
        with open("update_new.exe", "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        # Ersetzen der aktuellen .exe
        current_exe = sys.executable
        backup_exe = current_exe + ".old"
        os.rename(current_exe, backup_exe)
        os.rename("update_new.exe", current_exe)

        messagebox.showinfo("Update", "Update erfolgreich abgeschlossen. Bitte starten Sie die Anwendung neu.")
        sys.exit()  # Beenden der aktuellen Instanz
    except Exception as e:
        messagebox.showerror("Update-Fehler", f"Fehler beim Update: {e}")
        print(f"Update-Fehler: {e}")
