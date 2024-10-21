
import os

icon_path = "./icons/Screenshot_1.png"
if not os.path.exists(icon_path):
    print("Icon-Datei nicht gefunden:", icon_path)
else:
    print("Icon-Datei gefunden:", icon_path)