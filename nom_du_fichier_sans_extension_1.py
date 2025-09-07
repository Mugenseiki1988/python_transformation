import os
import pandas as pd
import fnmatch
import tqdm

# 📂 Répertoires sources
pml_root = r"C:\Program Files (x86)\AVEVA\Everything3D2.10"
cs_root = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\transposition_dll_2.1"
output_path = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\ETUDE AVEVA UIC ETC\export_structure_fichiers.xlsx"

# 📄 Extensions cibles
pml_extensions = ["*.pmlfrm", "*.pmlobj", "*.pmlfnc"]
cs_extension = "*.cs"

# 📋 Résultats
results = []

# 🔍 Recherche fichiers PML
print("🔍 Recherche des fichiers PML...")
for root, _, files in tqdm.tqdm(os.walk(pml_root), desc="Analyse PML", unit="dir"):
    for ext in pml_extensions:
        for file in fnmatch.filter(files, ext):
            name_no_ext, extension = os.path.splitext(file)
            results.append({
                "nom du fichier sans extension": name_no_ext,
                "nom du fichier avec extension": file,
                "extension du fichier": extension,
                "namespace": "",
                "DLL": ""
            })

# 🔍 Recherche fichiers C#
print("\n🔍 Recherche des fichiers .cs...")
for root, _, files in tqdm.tqdm(os.walk(cs_root), desc="Analyse C#", unit="dir"):
    for file in fnmatch.filter(files, cs_extension):
        name_no_ext, extension = os.path.splitext(file)
        relative_path = os.path.relpath(root, cs_root)
        path_parts = relative_path.split(os.sep)

        # Namespace = dernier dossier, DLL = dossier parent du namespace
        namespace = path_parts[-1] if len(path_parts) >= 1 else ""
        dll = path_parts[-2] if len(path_parts) >= 2 else ""

        results.append({
            "nom du fichier sans extension": name_no_ext,
            "extension du fichier": extension,
            "DLL": dll,
            "namespace": namespace,
            "nom du fichier avec extension": file,
        })

# 💾 Export Excel
df = pd.DataFrame(results)
df.to_excel(output_path, index=False)
print(f"\n✅ Export terminé : {output_path}")
