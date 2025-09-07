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

output_file = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\Decodeur langage PML1_2\\container_pmlnetcontrol.xlsx"

# --- Activation de la lecture depuis un fichier texte ---
use_txt_term_list = False
txt_term_file = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\liste_dll.txt"

# --- Activation des options ---
use_exact_terms = True
use_partial_terms = False

use_exclam_prefix_exact = False    # Active la recherche de !!term sur exact_terms
use_underscore_prefix_exact = False  # Active la recherche de _term sur exact_terms

use_exclam_prefix_partial = False     # Active la recherche de !!term sur partial_terms
use_underscore_prefix_partial = False  # Active la recherche de _term sur partial_terms

partial_terms_start_only = False  # Active la détection des mots commençant par le terme
exclude_partial_exceptions = False  # Exclut les cas comme CALLABLE, CALLBACK

# --- Activation ou non du traitement spécial des .pmlcmd ---
use_pmlcmd_special_block_processing = False

# --- Listes internes de termes 
exact_terms = [
    "container",
    "PmlNetControl",
]
partial_terms = ["!!CD", "_CD"]
exclude_words = ["CALLABLE", "CALLBACK"] if exclude_partial_exceptions else []

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

            line_lower = line.lower()

            # --- EXACT TERMS ---
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

                    prefixes = []
                    if use_exclam_prefix_exact:
                        prefixes.append(f"!!{term.lower()}")
                    if use_underscore_prefix_exact:
                        prefixes.append(f"_{term.lower()}")

                    for prefixed in prefixes:
                        if re.search(rf"(^|\W){re.escape(prefixed)}($|\W)", line_lower):
                            results.append({
                                "Chemin du fichier": file_path,
                                "Nom du fichier": os.path.basename(file_path),
                                "search_directory": base_dir,
                                "file_extension": ext,
                                "Terme détecté": term,
                                "Terme détecté entier": prefixed,
                                "Type de correspondance": "exact_terms",
                                "Ligne complète": line.rstrip('\n\r'),
                                "Numéro de ligne": line_number
                            })

            # --- PARTIAL TERMS ---
            if use_partial_terms:
                for term in partial_terms:
                    if term.startswith("!!") or term.startswith("_"):
                        base_pattern = rf"{re.escape(term.lower())}\w*"
                    else:
                        base_pattern = rf"\b{term.lower()}\w*" if partial_terms_start_only else rf"{term.lower()}\w*"

                    regex_base = re.compile(base_pattern, re.IGNORECASE)
                    matches = regex_base.findall(line)
                    for match_word in matches:
                        if match_word.upper() not in exclude_words:
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

                    patterns = []
                    if use_exclam_prefix_partial and term.startswith("!!"):
                        patterns.append(rf"{re.escape(term.lower())}\w*")
                    if use_underscore_prefix_partial:
                        patterns.append(rf"_{term.lower()}\w*")

                    for pat in patterns:
                        regex = re.compile(pat, re.IGNORECASE)
                        matches = regex.findall(line)
                        for match_word in matches:
                            if match_word.upper() not in exclude_words:
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

            # --- TXT TERMS ---
            for term in txt_terms:
                if term.lower() in line.lower():
                    results.append({
                        "Chemin du fichier": file_path,
                        "Nom du fichier": os.path.basename(file_path),
                        "search_directory": base_dir,
                        "file_extension": ext,
                        "Terme détecté": term,
                        "Terme détecté entier": term,
                        "Type de correspondance": "txt_terms",
                        "Ligne complète": line.rstrip('\n\r'),
                        "Numéro de ligne": line_number
                    })

    except Exception as e:
        print(f"Erreur lors de l'analyse de {file_path} : {e}")

    # 2/ === Bloc de traitement des fichiers .pmlcmd ===
    if use_pmlcmd_special_block_processing and ext.lower() == ".pmlcmd":
        in_method = False
        current_block = []
        start_line_num = None

        for line_number, line in enumerate(lines, start=1):
            stripped = line.strip()

            if stripped.lower().startswith("define method"):
                in_method = True
                current_block = [(line_number, line.rstrip('\n\r'))]
                start_line_num = line_number

            elif stripped.lower().startswith("endmethod") and in_method:
                current_block.append((line_number, line.rstrip('\n\r')))

                for lineno, content in current_block:
                    if not content.strip():
                        continue

                    detected_term = ""
                    detected_exact = ""
                    match_type = ""

                    lowered_line = content.lower()

                    if use_exact_terms:
                        for term in exact_terms:
                            if re.search(rf"(^|\W){re.escape(term.lower())}($|\W)", lowered_line):
                                detected_term = term
                                detected_exact = term
                                match_type = "exact_terms"
                                break

                    if not detected_term and use_partial_terms:
                        for term in partial_terms:
                            pattern = rf"\b{term.lower()}\w*" if partial_terms_start_only else rf"{term.lower()}\w*"
                            matches = re.findall(pattern, lowered_line)
                            for m in matches:
                                if m.upper() not in exclude_words:
                                    detected_term = term
                                    detected_exact = m
                                    match_type = "partial_terms"
                                    break
                            if detected_term:
                                break

                    results.append({
                        "Chemin du fichier": file_path,
                        "Nom du fichier": os.path.basename(file_path),
                        "search_directory": base_dir,
                        "file_extension": ext,
                        "Terme détecté": detected_term,
                        "Terme détecté entier": detected_exact,
                        "Type de correspondance": match_type,
                        "Ligne complète": content,
                        "Numéro de ligne": lineno
                    })

                in_method = False
                current_block = []

            elif in_method:
                if stripped:
                    current_block.append((line_number, line.rstrip('\n\r')))

# --- Étape finale : Ajout de la colonne "Type usage" et export Excel ---
if results:
    import pandas as pd
    import re

    df = pd.DataFrame(results)

    def detect_type_usage(row):
        term_entier = str(row["Terme détecté entier"]).lower()
        line = str(row["Ligne complète"]).lower()

        patterns = {
            "$*": rf"(?<!\w)\$\*{re.escape(term_entier)}(?!\w)",
            "$!!": rf"(?<!\w)\$!!{re.escape(term_entier)}(?!\w)",
            "!!": rf"(?<!\w)!!{re.escape(term_entier)}(?!\w)",
            "!": rf"(?<!\w)!{re.escape(term_entier)}(?!\w)",
            "_": rf"(?<!\w)_{re.escape(term_entier)}(?!\w)",
            ".": rf"(?<!\w)\.{re.escape(term_entier)}(?!\w)",  # ✅ AJOUT CORRECTEMENT PRIS EN COMPTE
        }

        for prefix, pattern in patterns.items():
            if re.search(pattern, line):
                return prefix
        return ""

    df["Type usage"] = df.apply(detect_type_usage, axis=1)

    # Réorganisation des colonnes
    ordered_columns = [
        "Chemin du fichier",
        "Nom du fichier",
        "search_directory",
        "file_extension",
        "Terme détecté",
        "Terme détecté entier",
        "Type usage",
        "Type de correspondance",
        "Ligne complète",
        "Numéro de ligne"
    ]
    df = df[ordered_columns]

    # Correction spécifique pour partial_terms : réécriture explicite du préfixe détecté dans "Terme détecté entier"
    def correct_type_usage_for_partial(row):
        if row["Type de correspondance"] != "partial_terms" or row["Type usage"]:
            return row["Type usage"]  # ne modifie rien si pas partial_terms ou déjà rempli

        term_detecte = row["Terme détecté"].strip()
        terme_entier = row["Terme détecté entier"].strip()

        if term_detecte.startswith("!!") and terme_entier.startswith("!!"):
            return "!!"
        elif term_detecte.startswith("_") and terme_entier.startswith("_"):
            return "_"
        return ""  # pas de correction possible

    df["Type usage"] = df.apply(correct_type_usage_for_partial, axis=1)

    # Suppression des doublons pour les lignes "exact_terms" avec Type usage vide
    mask_exact_terms = (df["Type de correspondance"] == "exact_terms") & (df["Type usage"] == "")
    
    # Colonnes à concaténer pour identifier les doublons
    dedup_cols = [
        "Chemin du fichier",
        "Nom du fichier",
        "search_directory",
        "file_extension",
        "Terme détecté",
        "Terme détecté entier",
        "Type de correspondance",
        "Ligne complète",
        "Numéro de ligne"
    ]
    
    # Création d'une clé de comparaison
    df["_dedup_key"] = df[dedup_cols].astype(str).agg("||".join, axis=1)

    # Supprimer les doublons dans ce sous-ensemble
    df = pd.concat([
        df[~mask_exact_terms],
        df[mask_exact_terms].drop_duplicates(subset="_dedup_key")
    ], ignore_index=True)

    # Nettoyage de la colonne temporaire
    df.drop(columns=["_dedup_key"], inplace=True)

    # Export Excel (en une ou plusieurs parties si besoin)
    max_rows = 1_048_576
    total_rows = len(df)

    if total_rows <= max_rows:
        df.to_excel(output_file, index=False)
        print(f"\n✅ Fichier exporté : {output_file} ({total_rows} lignes, avec Type usage)")
    else:
        num_parts = (total_rows // max_rows) + (1 if total_rows % max_rows > 0 else 0)
        for i in range(num_parts):
            start = i * max_rows
            end = min(start + max_rows, total_rows)
            df_part = df.iloc[start:end]
            part_file = output_file.replace(".xlsx", f"_part{i+1}.xlsx")
            df_part.to_excel(part_file, index=False)
            print(f"✅ Partie {i+1} exportée : {part_file} ({len(df_part)} lignes)")

else:
    print("\n❌ Aucun terme détecté.")