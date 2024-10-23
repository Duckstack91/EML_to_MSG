import os
import email
from email import policy
from email.parser import BytesParser
import win32com.client
import re
import traceback

def sanitize_filename(filename):
    """Bereinigt den Dateinamen, um Probleme mit Sonderzeichen zu vermeiden."""
    return re.sub(r'[^a-zA-Z0-9äöüÄÖÜß_\-\.]', '_', filename)

def create_output_folder(output_dir):
    """Erstellt den Ordner 'Eml_Konverter' im Zielverzeichnis, falls er nicht bereits innerhalb von 'Eml_Konverter' liegt."""
    try:
        if 'Eml_Konverter' in os.path.basename(output_dir):
            print(f"Der Ordner 'Eml_Konverter' existiert bereits im Pfad: {output_dir}")
            return output_dir
        output_folder = os.path.join(output_dir, 'Eml_Konverter')
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"Ordner 'Eml_Konverter' erstellt: {output_folder}")
        else:
            print(f"Ordner 'Eml_Konverter' existiert bereits: {output_folder}")
        return output_folder
    except Exception as e:
        print(f"Fehler beim Erstellen des Ordners: {e}")
        raise

def eml_to_msg(eml_file, output_dir, prefix, is_attachment=False):
    """Konvertiert eine .eml-Datei in .msg und speichert Anhänge"""
    try:
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
        eml_filename = os.path.splitext(os.path.basename(eml_file))[0]
        sanitized_eml_filename = sanitize_filename(eml_filename)

        # Benenne die Haupt-E-Mail mit "_01_" im Präfix, Anhang weiter zählen
        if not is_attachment:
            msg_filename = f"{prefix}01_{sanitized_eml_filename}.msg"
        else:
            msg_filename = f"{prefix}{sanitized_eml_filename}.msg"

        # Speicherort auf 'Eml_Konverter' festlegen
        output_folder = create_output_folder(output_dir)
        msg_file = os.path.join(output_folder, msg_filename)

        try:
            outlook_msg.SaveAs(msg_file)
            print(f"Erfolgreich konvertiert: {eml_file} zu {msg_file}")
        except Exception as e:
            print(f"Fehler beim Speichern der Nachricht: {e}")
            traceback.print_exc()
            return

        return output_folder

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten bei der Verarbeitung von {eml_file}: {e}")
        traceback.print_exc()

def process_attachments(msg, output_folder, prefix, start_counter):
    """Verarbeitet die Anhänge und nummeriert sie korrekt."""
    attachment_counter = start_counter

    try:
        for part in msg.iter_attachments():
            filename = part.get_filename()

            if filename:
                sanitized_filename = sanitize_filename(filename)
                attachment_filename = f"{prefix}{str(attachment_counter).zfill(2)}_{sanitized_filename}"
                attachment_path = os.path.join(output_folder, attachment_filename)

                # Falls der Anhang eine .eml-Datei ist, diese ebenfalls als .msg speichern
                if filename.endswith('.eml'):
                    with open(attachment_path, 'wb') as af:
                        af.write(part.get_payload(decode=True))
                    print(f"Verschachtelte EML-Datei gefunden: {attachment_path}")
                    attachment_counter += 1
                    eml_to_msg(attachment_path, output_folder, prefix, is_attachment=True)
                else:
                    # Speichere reguläre Anhänge
                    with open(attachment_path, 'wb') as af:
                        af.write(part.get_payload(decode=True))
                    print(f"Anhang gespeichert: {attachment_path}")

                attachment_counter += 1

    except Exception as e:
        print(f"Fehler beim Speichern der Anhänge: {e}")
        traceback.print_exc()

def process_directory(eml_directory):
    """Verarbeitet alle EML-Dateien im angegebenen Verzeichnis."""
    counter = 1
    for root, dirs, files in os.walk(eml_directory):
        files.sort()
        for file in files:
            prefix = f"{str(counter).zfill(4)}_"
            file_path = os.path.join(root, file)
            if file.endswith('.eml'):
                output_folder = eml_to_msg(file_path, root, prefix)
                if output_folder:
                    # Verarbeite die Anhänge ab der Nummer 02
                    with open(file_path, 'rb') as f:
                        msg = BytesParser(policy=policy.default).parse(f)
                        process_attachments(msg, output_folder, prefix, start_counter=2)
            counter += 1
