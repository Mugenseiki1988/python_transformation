import os
import re
import fnmatch
import pandas as pd
import tqdm

# --- Configuration ---
search_directories = [
    r"C:\\Program Files (x86)\\AVEVA\\Everything3D2.10",
    r"D:\\E3D.2.1\\MEIUI",
    r"D:\\E3D.2.1\\MEILIB"
]

file_extensions = ["*.pmlfrm", "*.pmlobj", "*.pmlcmd", "*.pmlfnc", "*.pmlmac", ""]  # Inclut les fichiers sans extension

output_file = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\ultime_resultats_recherche_pml_1.xlsx"

# --- Activation de la lecture depuis un fichier texte ---
use_txt_term_list = True
txt_term_file = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\liste_dll.txt"

# --- Listes internes de termes à détecter ---
exact_terms = [
    "import", "namespace"
]

partial_terms = ["dll", "browser"]

# --- Lecture des termes depuis un fichier texte ---
txt_terms = []
if use_txt_term_list and os.path.isfile(txt_term_file):
    with open(txt_term_file, "r", encoding="utf-8", errors="ignore") as f:
        txt_terms = [line.strip() for line in f if line.strip()]

# --- Collecte des fichiers ---
pml_files = []
for base_dir in search_directories:
    for root, _, files in os.walk(base_dir):
        for file in files:
            file_path = os.path.join(root, file)
            ext = os.path.splitext(file)[1]
            if ext and any(fnmatch.fnmatch(file.lower(), pattern.lower()) for pattern in file_extensions if pattern):
                pml_files.append((file_path, base_dir, ext))
            elif not ext and "" in file_extensions:
                pml_files.append((file_path, base_dir, ""))

# --- Analyse des fichiers ---
results = []

for file_path, base_dir, ext in tqdm.tqdm(pml_files, desc="Analyse PML", unit="fichier"):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()

        for line_number, line in enumerate(lines, start=1):
            line_clean = line.strip()
            if not line_clean or re.match(r"^[ \t]*-", line_clean):
                continue

            # Recherche exacte (insensible à la casse)
            for term in exact_terms:
                if term.lower() in line.lower():
                    results.append({
                        "Chemin du fichier": file_path,
                        "Nom du fichier": os.path.basename(file_path),
                        "search_directory": base_dir,
                        "file_extension": ext,
                        "Terme détecté": term,
                        "Type de correspondance": "exact_terms",
                        "Ligne complète": line.rstrip('\n\r'),
                        "Numéro de ligne": line_number
                    })

            # Recherche partielle
            for term in partial_terms:
                if term.lower() in line.lower():
                    results.append({
                        "Chemin du fichier": file_path,
                        "Nom du fichier": os.path.basename(file_path),
                        "search_directory": base_dir,
                        "file_extension": ext,
                        "Terme détecté": term,
                        "Type de correspondance": "partial_terms",
                        "Ligne complète": line.rstrip('\n\r'),
                        "Numéro de ligne": line_number
                    })

            # Recherche à partir du fichier texte
            for term in txt_terms:
                if term.lower() in line.lower():
                    results.append({
                        "Chemin du fichier": file_path,
                        "Nom du fichier": os.path.basename(file_path),
                        "search_directory": base_dir,
                        "file_extension": ext,
                        "Terme détecté": term,
                        "Type de correspondance": "txt_terms",
                        "Ligne complète": line.rstrip('\n\r'),
                        "Numéro de ligne": line_number
                    })

    except Exception as e:
        print(f"Erreur lors de l'analyse de {file_path} : {e}")

# --- Export ---
if results:
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"\n✅ Fichier exporté : {output_file}")
else:
    print("\n❌ Aucun terme détecté.")
