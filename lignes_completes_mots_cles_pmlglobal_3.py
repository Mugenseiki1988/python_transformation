import os
import pandas as pd
import fnmatch
import tqdm
import re

# 📂 Définition des chemins
search_directory = r"C:\Program Files (x86)\AVEVA\Everything3D2.10"
output_file = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\ETUDE AVEVA UIC ETC\recherche_lignes_completes_mots_cles_container_pmlcontrol_3.xlsx"

# 🔍 Terme à rechercher (dans tout mot, insensible à la casse)
search_terms = ["container", "pmlcontrol"]

# 📋 Stockage des résultats
results = []

# 📂 Extensions de fichiers à scanner
file_extensions = ["*.pmlfrm", "*.pmlobj", "*.pmlcmd", "*.pmlfnc", "*.mac", "*.pmlmac"]

# 📂 Recherche récursive des fichiers
pml_files = []
for root, _, files in os.walk(search_directory):
    for file in files:
        if any(fnmatch.fnmatch(file.lower(), ext.lower()) for ext in file_extensions):
            pml_files.append(os.path.join(root, file))

# 📊 Analyse des fichiers avec barre de progression
for file_path in tqdm.tqdm(pml_files, desc="🔎 Analyse des fichiers PML", unit="fichier"):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
            for line_number, line in enumerate(lines, start=1):
                for term in search_terms:
                    if term.lower() in line.lower():
                        # Cherche toute méthode appelée avec un mot contenant "drawlist"
                        pattern = r'(\w*' + re.escape(term) + r'\w*)\.([a-zA-Z_]\w*)'
                        matches = re.findall(pattern, line, re.IGNORECASE)
                        for full_var, method in matches:
                            relative_path = os.path.relpath(file_path, search_directory)
                            results.append({
                                "Chemin du fichier": relative_path,
                                "Nom du fichier": os.path.basename(file_path),
                                "Terme détecté": term,
                                "Ligne complète": line.strip(),
                                "Numéro de ligne": line_number,
                                "Méthode appelée": f".{method}"
                            })
    except Exception as e:
        print(f"⚠️ Erreur lors de l'analyse de {file_path} : {e}")

# 📤 Exportation Excel
if results:
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"\n✅ Résultats enregistrés dans : {output_file}")
else:
    print("\n❌ Aucun terme trouvé dans les fichiers analysés.")