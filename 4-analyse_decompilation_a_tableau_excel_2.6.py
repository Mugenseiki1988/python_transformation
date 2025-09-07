import os
import json
import csv
import re
import hashlib
import xml.etree.ElementTree as ET
from tqdm import tqdm  # Barre de progression

# Dossier contenant les fichiers après décompilation
decompilation_folder = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\transposition_dll"

# Dossier de sortie général
output_root_folder = os.path.join(decompilation_folder, "analyse_decompilation")
os.makedirs(output_root_folder, exist_ok=True)  # Création du dossier principal

# Fonction pour tronquer un chemin trop long et le remplacer par un hash
def truncate_path(path, max_length=100):
    """Réduit la longueur d'un chemin trop long en le remplaçant par un hash"""
    if len(path) > max_length:
        hash_part = hashlib.md5(path.encode()).hexdigest()[:10]  # Hash de 10 caractères
        path = path[:max_length - 11] + "_" + hash_part
    return path

# Expressions régulières pour extraire les informations des fichiers .cs
regex_namespace = re.compile(r'namespace\s+([\w.]+)')
regex_class = re.compile(r'class\s+([\w]+)')
regex_pmlnet = re.compile(r'\[PMLNetCallable\]')  # Détecte si une ligne contient PMLNetCallable
regex_method = re.compile(r'(\[PMLNetCallable\])?\s*(public|private)?\s*([\w<>\[\]]+)\s+([\w]+)\((.*?)\)')
regex_property = re.compile(r'(public|private)?\s*([\w<>[\]]+)\s+([\w]+)\s*{.*?get;.*?set;.*?}')
regex_variable = re.compile(r'(public|private)?\s*([\w<>[\]]+)\s+([\w]+);')
regex_event = re.compile(r'public\s+event\s+([\w<>[\]]+)\s+([\w]+);')

# Récupération de la liste des fichiers à analyser par répertoire
all_files_by_directory = {}

for root, _, files in os.walk(decompilation_folder):
    if files:
        relative_path = os.path.relpath(root, decompilation_folder)
        truncated_relative_path = truncate_path(relative_path)  # Tronquer si trop long
        output_folder = os.path.join(output_root_folder, truncated_relative_path)  # Dossier spécifique pour ce répertoire
        os.makedirs(output_folder, exist_ok=True)  # Création du dossier de sortie s'il n'existe pas
        all_files_by_directory[root] = {"files": [os.path.join(root, file) for file in files], "output_folder": output_folder}

print(f"📂 Nombre total de répertoires analysés : {len(all_files_by_directory)}")

# Parcours des répertoires et analyse des fichiers
for directory, data in tqdm(all_files_by_directory.items(), desc="🔍 Analyse des répertoires", unit="dir"):
    files = data["files"]
    output_folder = data["output_folder"]
    
    data_list = []
    unreadable_files = []
    index_beta = 1  # Indexation alternative

    for file_path in tqdm(files, desc=f"📄 Analyse des fichiers sous {directory}", unit="file"):
        file = os.path.basename(file_path)
        file_extension = os.path.splitext(file)[1]
        relative_path = os.path.relpath(os.path.dirname(file_path), decompilation_folder)
        truncated_relative_path = truncate_path(relative_path)  # Tronquer si trop long

        # Définition des indexations et niveaux
        niveau_beta = len(relative_path.split(os.sep))
        pair_impair_beta = "Pair" if index_beta % 2 == 0 else "Impair"
        index_niveau_beta = f"{index_beta}-{niveau_beta}"

        try:
            # Analyser les fichiers C# (.cs)
            if file.endswith(".cs"):
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.readlines()

                full_text = "".join(content)

                namespaces = regex_namespace.findall(full_text) or ["N/A"]
                classes = regex_class.findall(full_text) or ["N/A"]
                methods = regex_method.findall(full_text)
                properties = regex_property.findall(full_text)
                variables = regex_variable.findall(full_text)
                events = regex_event.findall(full_text)

                # Ajouter chaque méthode ligne par ligne
                for method in methods:
                    data_list.append({
                        "Index_beta": index_beta,
                        "Pair_Impair_beta": pair_impair_beta,
                        "Niveau_beta": niveau_beta,
                        "Index-Niveau_beta": index_niveau_beta,
                        "Repertoire": truncated_relative_path,
                        "Fichier_Source": file,
                        "Extension": file_extension,
                        "Namespace": ", ".join(namespaces),
                        "Classe": ", ".join(classes),
                        "Methode": method[3],
                        "Public/Private": method[1] or "N/A",
                        "Modificateur": method[2],
                        "Paramètres": method[4],
                        "PMLNetCallable": "Oui" if method[0] else "Non",
                        "Propriete": "N/A",
                        "Evenement": "N/A"
                    })

                # Ajouter chaque propriété ligne par ligne
                for prop in properties:
                    data_list.append({
                        "Index_beta": index_beta,
                        "Pair_Impair_beta": pair_impair_beta,
                        "Niveau_beta": niveau_beta,
                        "Index-Niveau_beta": index_niveau_beta,
                        "Repertoire": truncated_relative_path,
                        "Fichier_Source": file,
                        "Extension": file_extension,
                        "Namespace": ", ".join(namespaces),
                        "Classe": ", ".join(classes),
                        "Methode": "N/A",
                        "Public/Private": prop[0] or "N/A",
                        "Modificateur": prop[1],
                        "Propriete": prop[2],
                        "Evenement": "N/A",
                        "PMLNetCallable": "Non"
                    })

                # Ajouter chaque variable
                for var in variables:
                    data_list.append({
                        "Index_beta": index_beta,
                        "Pair_Impair_beta": pair_impair_beta,
                        "Niveau_beta": niveau_beta,
                        "Index-Niveau_beta": index_niveau_beta,
                        "Repertoire": truncated_relative_path,
                        "Fichier_Source": file,
                        "Extension": file_extension,
                        "Namespace": ", ".join(namespaces),
                        "Classe": ", ".join(classes),
                        "Methode": "N/A",
                        "Public/Private": var[0] or "N/A",
                        "Modificateur": var[1],
                        "Propriete": var[2],
                        "Evenement": "N/A",
                        "PMLNetCallable": "Non"
                    })

            index_beta += 1  # Incrémentation de l'index beta

        except Exception:
            unreadable_files.append(file_path)

    # Sauvegarde en JSON et CSV pour ce répertoire
    if data_list:
        output_json = os.path.join(output_folder, "decompilation_analysis.json")
        output_csv = os.path.join(output_folder, "decompilation_analysis.csv")

        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(data_list, f, indent=4, ensure_ascii=False)

        with open(output_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=data_list[0].keys())
            writer.writeheader()
            writer.writerows(data_list)

print("\n✅ Analyse terminée. Les fichiers JSON et CSV sont générés par répertoire.")