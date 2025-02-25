import importlib.util
import os

# `version.py` dynamisch importieren
spec = importlib.util.spec_from_file_location("version", "version.py")
version_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(version_module)

# Version aus `version.py`
version = version_module.VERSION

# 📝 Version nur als Zahl für GitHub speichern
github_version_path = "src/version.txt"
os.makedirs(os.path.dirname(github_version_path), exist_ok=True)
with open(github_version_path, "w", encoding="utf-8") as f:
    f.write(version)  # Speichert nur `1.1.2`

# 📦 Version für PyInstaller als `version.txt`
pyinstaller_version_path = "version.txt"
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
with open(pyinstaller_version_path, "w", encoding="utf-8") as f:
    f.write(version_info)

print(f"✅ Version {version} erfolgreich geschrieben:")
print(f"  - PyInstaller: {pyinstaller_version_path}")
print(f"  - GitHub: {github_version_path}")
