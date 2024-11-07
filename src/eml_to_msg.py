import os
import email
from email import policy
from email.parser import BytesParser
import win32com.client
import re
import traceback


def sanitize_filename(filename):
    return re.sub(r'[^a-zA-Z0-9äöüÄÖÜß_\-\.]', '_', filename)


def eml_to_msg(eml_file, output_dir, prefix, is_attachment=False):
    try:
        with open(eml_file, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook_msg = outlook.CreateItem(0)
        except Exception as e:
            print(f"Fehler bei der Initialisierung von Outlook: {e}")
            traceback.print_exc()
            return

        try:
            subject = msg.get('Subject', 'Kein Betreff')
            sanitized_subject = sanitize_filename(subject)
            outlook_msg.Subject = subject
            outlook_msg.Body = msg.get_body(preferencelist=('plain')).get_content() if msg.get_body(
                preferencelist=('plain')) else ''
            outlook_msg.HTMLBody = msg.get_body(preferencelist=('html')).get_content() if msg.get_body(
                preferencelist=('html')) else ''
            outlook_msg.Sender = msg.get('From', '')
            outlook_msg.To = msg.get('To', '')
            outlook_msg.CC = msg.get('CC', '')
        except Exception as e:
            print(f"Fehler beim Setzen der E-Mail-Eigenschaften: {e}")
            traceback.print_exc()
            return

        eml_filename = os.path.splitext(os.path.basename(eml_file))[0]
        sanitized_eml_filename = sanitize_filename(eml_filename)
        msg_filename = f"{prefix}01_{sanitized_eml_filename}.msg" if not is_attachment else f"{prefix}{sanitized_eml_filename}.msg"

        msg_file = os.path.join(output_dir, msg_filename)

        try:
            outlook_msg.SaveAs(msg_file)
            print(f"Erfolgreich konvertiert: {eml_file} zu {msg_file}")
        except Exception as e:
            print(f"Fehler beim Speichern der Nachricht: {e}")
            traceback.print_exc()
            return

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten bei der Verarbeitung von {eml_file}: {e}")
        traceback.print_exc()


def process_attachments(msg, output_folder, prefix, start_counter):
    attachment_counter = start_counter
    try:
        for part in msg.iter_attachments():
            filename = part.get_filename()
            if filename:
                sanitized_filename = sanitize_filename(filename)
                attachment_filename = f"{prefix}{str(attachment_counter).zfill(2)}_{sanitized_filename}"
                attachment_path = os.path.join(output_folder, attachment_filename)

                if filename.endswith('.eml'):
                    with open(attachment_path, 'wb') as af:
                        af.write(part.get_payload(decode=True))
                    eml_to_msg(attachment_path, output_folder, prefix, is_attachment=True)
                else:
                    with open(attachment_path, 'wb') as af:
                        af.write(part.get_payload(decode=True))
                attachment_counter += 1
    except Exception as e:
        print(f"Fehler beim Speichern der Anhänge: {e}")
        traceback.print_exc()


def process_directory(eml_directory, output_directory):
    counter = 1
    for root, dirs, files in os.walk(eml_directory):
        files.sort()
        for file in files:
            prefix = f"{str(counter).zfill(4)}_"
            file_path = os.path.join(root, file)
            if file.endswith('.eml'):
                eml_to_msg(file_path, output_directory, prefix)
                with open(file_path, 'rb') as f:
                    msg = BytesParser(policy=policy.default).parse(f)
                    process_attachments(msg, output_directory, prefix, start_counter=2)
            counter += 1
