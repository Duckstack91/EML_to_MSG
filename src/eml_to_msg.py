import os
import email
from email import policy
from email.parser import BytesParser
import win32com.client
import re
import traceback

def sanitize_filename(filename):
    """Bereinigt den Dateinamen, um Probleme mit Sonderzeichen zu vermeiden."""
    return re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename)

def eml_to_msg(eml_file, output_dir, prefix):
    try:
        # Parse die .eml-Datei
        with open(eml_file, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        # Erstelle ein Outlook-Nachrichtenobjekt
        try:
            print("Versuche, Outlook zu initialisieren...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook_msg = outlook.CreateItem(0)
            print("Outlook erfolgreich initialisiert.")
        except Exception as e:
            print(f"Fehler bei der Initialisierung von Outlook: {e}")
            traceback.print_exc()
            return

        # Setze die E-Mail-Eigenschaften
        try:
            subject = msg.get('Subject', 'Kein Betreff')
            sanitized_subject = sanitize_filename(subject)
            outlook_msg.Subject = subject
            outlook_msg.Body = msg.get_body(preferencelist=('plain')).get_content() if msg.get_body(preferencelist=('plain')) else ''
            outlook_msg.HTMLBody = msg.get_body(preferencelist=('html')).get_content() if msg.get_body(preferencelist=('html')) else ''
            outlook_msg.Sender = msg.get('From', '')
            outlook_msg.To = msg.get('To', '')
            outlook_msg.CC = msg.get('CC', '')
        except Exception as e:
            print(f"Fehler beim Setzen der E-Mail-Eigenschaften: {e}")
            traceback.print_exc()
            return

        # Bereinige den .eml-Dateinamen für das Speichern der .msg-Datei
        eml_filename = os.path.splitext(os.path.basename(eml_file))[0]  # Entferne die Dateiendung ".eml"
        sanitized_eml_filename = sanitize_filename(eml_filename)

        # Speichere die Nachricht als .msg mit Präfix und bereinigtem .eml-Namen
        try:
            msg_filename = f"{prefix}01_{sanitized_eml_filename}.msg"
            msg_file = os.path.join(output_dir, msg_filename)
            outlook_msg.SaveAs(msg_file)
            print(f"Erfolgreich konvertiert: {eml_file} zu {msg_file}")
        except Exception as e:
            print(f"Fehler beim Speichern der Nachricht: {e}")
            traceback.print_exc()
            return

        # Speichere Anhänge
        try:
            attachment_counter = 2  # Beginne bei 02 für Anhänge
            for part in msg.iter_attachments():
                filename = part.get_filename()

                if filename:  # Überprüfe, ob der Anhang einen Dateinamen hat
                    sanitized_filename = sanitize_filename(filename)
                    attachment_filename = f"{prefix}{str(attachment_counter).zfill(2)}_{sanitized_filename}"
                    attachment_path = os.path.join(output_dir, attachment_filename)

                    # Falls der Anhang eine .eml-Datei ist, diese ebenfalls als .msg speichern
                    if filename.endswith('.eml'):
                        with open(attachment_path, 'wb') as af:
                            af.write(part.get_payload(decode=True))
                        # Jetzt die verschachtelte .eml-Datei konvertieren
                        print(f"Verschachtelte EML-Datei gefunden: {attachment_path}")
                        eml_to_msg(attachment_path, output_dir, f"{prefix}{str(attachment_counter).zfill(2)}_")
                    else:
                        with open(attachment_path, 'wb') as af:
                            af.write(part.get_payload(decode=True))
                        print(f"Anhang gespeichert: {attachment_path}")

                    attachment_counter += 1

        except Exception as e:
            print(f"Fehler beim Speichern der Anhänge: {e}")
            traceback.print_exc()
            return

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten bei der Verarbeitung von {eml_file}: {e}")
        traceback.print_exc()

def process_directory(eml_directory):
    """Verarbeitet alle EML-Dateien im angegebenen Verzeichnis."""
    counter = 1
    for root, dirs, files in os.walk(eml_directory):
        files.sort()  # Sortiere die Dateien alphabetisch
        for file in files:
            prefix = f"{str(counter).zfill(4)}_"  # Präfix für jede E-Mail und deren Anhänge
            file_path = os.path.join(root, file)
            if file.endswith('.eml'):
                eml_to_msg(file_path, root, prefix)
            else:
                # Falls es keine .eml-Datei ist, nur den Dateinamen ändern
                sanitized_filename = sanitize_filename(file)
                new_filename = f"{prefix}{sanitized_filename}"
                new_file_path = os.path.join(root, new_filename)
                os.rename(file_path, new_file_path)
                print(f"Datei umbenannt: {file_path} zu {new_file_path}")
            counter += 1
