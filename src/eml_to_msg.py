import os
import email
from email import policy
from email.parser import BytesParser
import win32com.client
import re
import traceback
from idlelib.tooltip import Hovertip

def sanitize_filename(filename):
    """Sanitize filename to prevent any issues with file naming."""
    return re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename)
#test
def eml_to_msg(eml_file, output_dir, prefix):
    try:
        # Parse the .eml file
        with open(eml_file, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        # Create an Outlook message object
        try:
            print("Attempting to initialize Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook_msg = outlook.CreateItem(0)
            print("Outlook initialized successfully.")
        except Exception as e:
            print(f"Error initializing Outlook: {e}")
            traceback.print_exc()
            return

        # Set email properties
        try:
            subject = msg.get('Subject', 'No Subject')
            sanitized_subject = sanitize_filename(subject)
            outlook_msg.Subject = subject
            outlook_msg.Body = msg.get_body(preferencelist=('plain')).get_content() if msg.get_body(preferencelist=('plain')) else ''
            outlook_msg.HTMLBody = msg.get_body(preferencelist=('html')).get_content() if msg.get_body(preferencelist=('html')) else ''
            outlook_msg.Sender = msg.get('From', '')
            outlook_msg.To = msg.get('To', '')
            outlook_msg.CC = msg.get('CC', '')
        except Exception as e:
            print(f"Error setting email properties: {e}")
            traceback.print_exc()
            return

        # Save the message
        try:
            msg_filename = f"{prefix}01.msg"
            msg_file = os.path.join(output_dir, msg_filename)
            outlook_msg.SaveAs(msg_file)
            print(f"Successfully converted {eml_file} to {msg_file}")
        except Exception as e:
            print(f"Error saving the message: {e}")
            traceback.print_exc()
            return

        # Anhänge speichern
        try:
            attachment_counter = 2
            for part in msg.iter_attachments():
                filename = part.get_filename()
                if filename:
                    sanitized_filename = sanitize_filename(filename)
                    attachment_filename = f"{prefix}{str(attachment_counter).zfill(2)}_{sanitized_filename}"
                    attachment_path = os.path.join(output_dir, attachment_filename)
                    with open(attachment_path, 'wb') as af:
                        af.write(part.get_payload(decode=True))
                    attachment_counter += 1
        except Exception as e:
            print(f"Error saving attachments: {e}")
            traceback.print_exc()
            return

    except Exception as e:
        print(f"An error occurred while processing {eml_file}: {e}")
        traceback.print_exc()

def process_directory(eml_directory):
    counter = 1
    for root, dirs, files in os.walk(eml_directory):
        # Sort files alphabetically
        files.sort()
        for file in files:
            prefix = f"{str(counter).zfill(4)}_"
            file_path = os.path.join(root, file)
            if file.endswith('.eml'):
                eml_to_msg(file_path, root, prefix)
            else:
                # Simply rename the file with the prefix
                sanitized_filename = sanitize_filename(file)
                new_filename = f"{prefix}{sanitized_filename}"
                new_file_path = os.path.join(root, new_filename)
                os.rename(file_path, new_file_path)
                print(f"Renamed {file_path} to {new_file_path}")
            counter += 1

if __name__ == "__main__":
    eml_directory = "C:/Program Files (x86)/Programmieren/EML_to_MSG/venv/eml_files"
    process_directory(eml_directory)
