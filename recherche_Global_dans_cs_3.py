import os
import pandas as pd
import fnmatch
import tqdm

# 📂 Définition du chemin de recherche
cs_search_directory = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\transposition_dll_2.1"
output_file = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\ETUDE AVEVA UIC ETC\recherche_TERMS_dans_CS_sType Description.xlsx"

# 🔍 Liste des termes à rechercher
search_terms = ["sType", "Description (Full Description)", "cattext", "dtxrtext", "Description"]

# 📋 Stockage des résultats
results = []

# 📂 Extensions de fichiers à scanner (uniquement .cs)
file_extensions = ["*.cs"]

# 📂 Fonction de recherche récursive
def scan_files(directory, file_extensions):
    scanned_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if any(fnmatch.fnmatch(file, ext) for ext in file_extensions):
                scanned_files.append(os.path.join(root, file))
    return scanned_files

# 🔍 Recherche des fichiers C#
cs_files = scan_files(cs_search_directory, file_extensions)

# 📊 Analyse des fichiers avec une barre de progression
total_files = len(cs_files)

print(f"🔍 Début de l'analyse de {total_files} fichiers .cs dans : {cs_search_directory}")

for file_path in tqdm.tqdm(cs_files, desc="🔎 Analyse des fichiers C#", unit="fichier"):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
            detected_terms = [term for term in search_terms if term in content]

            if detected_terms:
                relative_path = os.path.relpath(file_path, cs_search_directory)
                results.append({
                    "Source": "C#",
                    "Chemin du fichier": relative_path,
                    "Nom du fichier": os.path.basename(file_path),
                    "Termes détectés": ", ".join(detected_terms)
                })
    except Exception as e:
        print(f"⚠️ Erreur lors de l'analyse de {file_path} : {e}")

# 📤 Exportation des résultats en Excel
if results:
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"\n✅ Analyse terminée. Résultats enregistrés dans : {output_file}")
else:
    print("\n❌ Aucun fichier .cs ne contient les termes recherchés.")