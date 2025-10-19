import os
import re
import fnmatch
import pandas as pd
import tqdm

# --- Configuration ---
search_directories = [
#    r"C:\\Program Files (x86)\\AVEVA\\Everything3D2.10",
    r"D:\\E3D.2.1\\MEIUI",
    r"D:\\E3D.2.1\\MEILIB",
#    r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\PROJETS fichiers Y W N",
#    r"D:\ORF"
]

file_extensions = ["*.pmlfrm", "*.pmlobj", "*.pmlcmd", "*.pmlfnc", "*.pmlmac", "*.bat", ""]

output_file = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\Decodeur langage PML1_2\\requetage_UDA.xlsx"
nom_fichier_txt = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\Decodeur langage PML1_2\\0 - liste des noms de fichiers pml & UI\\Nom_fichier.txt"

# --- Activation des options ---
use_exact_terms = True
use_partial_terms = False
partial_terms_start_only = False
exclude_partial_exceptions = False

# --- Listes internes ---
exact_terms = ["UDA"]
partial_terms = ["isdpreviewiso"]
exclude_words_path = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\Decodeur langage PML1_2\\2 - externe - CAL\\exclude_words.txt"

# Lecture du fichier txt de noms de fichiers (en minuscules)
try:
    with open(nom_fichier_txt, "r", encoding="utf-8") as f:
        list_nom_fichiers_txt = [line.strip().lower() for line in f if line.strip()]
except:
    list_nom_fichiers_txt = []

# Chargement des mots à exclure
exclude_words = set()
if exclude_partial_exceptions and os.path.isfile(exclude_words_path):
    with open(exclude_words_path, "r", encoding="utf-8") as f:
        exclude_words = set(word.strip().lower() for word in f if word.strip())

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

# --- Analyse ---
results = []

for file_path, base_dir, ext in tqdm.tqdm(pml_files, desc="Analyse PML", unit="fichier"):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()

        for line_number, line in enumerate(lines, start=1):
            line_clean = line.strip()
            if not line_clean:
                continue
            if re.match(r"^[ \t]*[$#\-]|^[ \t]*synonym", line_clean, re.IGNORECASE):
                continue

            line_lower = line.lower()

            # --- Exact terms ---
            if use_exact_terms:
                for term in exact_terms:
                    if re.search(rf"(^|\W){re.escape(term.lower())}($|\W)", line_lower):
                        results.append({
                            "Chemin du fichier": file_path,
                            "Nom du fichier": os.path.basename(file_path),
                            "search_directory": base_dir,
                            "file_extension": ext,
                            "Terme détecté": term,
                            "Terme détecté entier": term,
                            "Type de correspondance": "exact_terms",
                            "Ligne complète": line.rstrip('\n\r'),
                            "Numéro de ligne": line_number
                        })

            # --- Partial terms ---
            if use_partial_terms:
                for term in partial_terms:
                    pattern = rf"\b{term.lower()}\w*" if partial_terms_start_only else rf"{term.lower()}\w*"
                    matches = re.findall(pattern, line_lower)
                    for match_word in matches:
                        if match_word.lower() not in exclude_words:
                            results.append({
                                "Chemin du fichier": file_path,
                                "Nom du fichier": os.path.basename(file_path),
                                "search_directory": base_dir,
                                "file_extension": ext,
                                "Terme détecté": term,
                                "Terme détecté entier": match_word,
                                "Type de correspondance": "partial_terms",
                                "Ligne complète": line.rstrip('\n\r'),
                                "Numéro de ligne": line_number
                            })

    except Exception as e:
        print(f"Erreur lors de l'analyse de {file_path} : {e}")

# --- Export ---
if results:
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"\n✅ Fichier exporté : {output_file} ({len(df)} lignes)")
else:
    print("\n❌ Aucun terme détecté.")