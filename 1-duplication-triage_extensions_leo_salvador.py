import os
import shutil
from tqdm import tqdm  # Bibliothèque pour afficher la barre de progression

# Répertoire de départ et de destination
source_directory = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\leo_salvador"
destination_directory = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\leo_salvador_extension_tri"

# Créer le répertoire de destination s'il n'existe pas
os.makedirs(destination_directory, exist_ok=True)

# Étape 1 : Évaluation du nombre de fichiers à traiter
all_files = []
for root, _, files in os.walk(source_directory):
    for file in files:
        all_files.append(os.path.join(root, file))

total_files = len(all_files)
print(f"Nombre total de fichiers détectés à dupliquer et trier : {total_files}")

# Étape 2 : Traitement des fichiers avec barre de progression
copied_files = 0  # Compteur de fichiers copiés

for file_path in tqdm(all_files, desc="Duplication et tri des fichiers", unit="fichier"):
    file_name = os.path.basename(file_path)
    file_extension = os.path.splitext(file_name)[-1].lower().strip('.')  # Récupérer l'extension

    if file_extension:  # Vérifier que l'extension existe
        ext_directory = os.path.join(destination_directory, file_extension.upper())

        # Créer le dossier pour l'extension s'il n'existe pas
        os.makedirs(ext_directory, exist_ok=True)

        # Copier le fichier dans le dossier correspondant
        shutil.copy(file_path, os.path.join(ext_directory, file_name))
        copied_files += 1

# Étape 3 : Affichage du résumé
print(f"\nDuplication et tri des fichiers terminés.")
print(f"Nombre total de fichiers dupliqués et triés : {copied_files} / {total_files}")