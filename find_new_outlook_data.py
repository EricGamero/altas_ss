import os
import glob

appdata = os.environ.get("LOCALAPPDATA", "")
outlook_dir = os.path.join(appdata, "Packages", "Microsoft.OutlookForWindows_8wekyb3d8bbwe")

print(f"Buscando archivos en: {outlook_dir}")

matches = []
for root, dirs, files in os.walk(outlook_dir):
    for file in files:
        if file.endswith((".db", ".sqlite", ".dat", ".ldb", ".log", ".indexeddb")):
            matches.append(os.path.join(root, file))

print(f"Total archivos de base de datos / caché encontrados: {len(matches)}")
for m in matches[:15]:
    print(" -", m)
