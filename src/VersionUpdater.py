import importlib.util
import subprocess
import os

# version.py dynamisch importieren
spec = importlib.util.spec_from_file_location("version", "version.py")
version_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(version_module)

# Version aus version.py
version = version_module.VERSION

# Version als Ressourcendatei erstellen
version_info = f"""
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({version.replace('.', ',')}, 0),
    prodvers=({version.replace('.', ',')}, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        '040904b0',
        [
        StringStruct('CompanyName', 'Stadt Esslingen am Neckar'),
        StringStruct('FileDescription', 'EML zu MSG Konverter'),
        StringStruct('FileVersion', '{version}.0'),
        StringStruct('InternalName', 'eml_to_msg'),
        StringStruct('LegalCopyright', 'Copyright © 2024 Sercan Yurdusever'),
        StringStruct('OriginalFilename', 'eml_to_msg.exe'),
        StringStruct('ProductName', 'EML to MSG Converter'),
        StringStruct('ProductVersion', '{version}.0')
        ]
      )
      ]
    ),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""

# Version.txt schreiben
with open("version.txt", "w", encoding="utf-8") as f:
    f.write(version_info)

subprocess.run([
    "pyinstaller",
    "--onefile",
    "--version-file=version.txt",
    "--name Eml_to_Msg.exe ",
    "--windowed",
    "gui.py"
])

print(f"Version {version} erfolgreich in version.txt geschrieben.")
