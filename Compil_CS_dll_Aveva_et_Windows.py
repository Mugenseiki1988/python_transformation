import os
from tqdm import tqdm

# 📁 Chemin de départ (répertoire à parcourir)
racine = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\transposition_dll_2.1\Aveva.ApplicationFramework.Presentation.Implementation"

# 📄 Fichier de sortie (dans le dossier parent "DLL decompilation")
fichier_sortie = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\Compil_CS_dll_Aveva.ApplicationFramework.Presentation.Implementation.txt"

# 🔍 Extensions de fichiers cibles
extensions_cibles = {".cs", ".csproj"}

# 📋 Liste des fichiers à traiter
fichiers_a_traiter = []
for dossier_courant, _, fichiers in os.walk(racine):
    for fichier in fichiers:
        extension = os.path.splitext(fichier)[1].lower()
        if extension in extensions_cibles:
            fichiers_a_traiter.append(os.path.join(dossier_courant, fichier))

# 🖊️ Écriture dans le fichier avec barre de progression
with open(fichier_sortie, "w", encoding="utf-8") as sortie:
    for chemin_fichier in tqdm(fichiers_a_traiter, desc="📦 Compilation des fichiers", unit="fichier"):
        try:
            with open(chemin_fichier, "r", encoding="utf-8", errors="ignore") as f:
                contenu = f.read()
        except Exception as e:
            contenu = f"[Erreur lors de la lecture du fichier {chemin_fichier} : {e}]"

        chemin_relatif = os.path.relpath(chemin_fichier, racine)
        sortie.write(f"{chemin_relatif}\n\n")
        sortie.write(contenu)
        sortie.write("\n" + "_"*120 + "\n\n")  # Séparateur visuel

print(f"\n✅ Compilation terminée. Fichier généré : {fichier_sortie}")