import os
from tqdm import tqdm

# 📁 Chemins
racine = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\transposition_dll_2.1"
output_txt_root = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\txt_transposition_dll_2.1"

# 📂 Extensions cibles
extensions_cibles = {".cs", ".csproj"}

# 📁 Créer le dossier de sortie s'il n'existe pas
os.makedirs(output_txt_root, exist_ok=True)

# 📌 Liste des sous-dossiers DLL à compiler
sous_dossiers = [d for d in os.listdir(racine) if os.path.isdir(os.path.join(racine, d))]

for dossier in tqdm(sous_dossiers, desc="📦 Compilation des DLL décompilées", unit="dll"):
    dossier_complet = os.path.join(racine, dossier)
    output_file = os.path.join(output_txt_root, f"{dossier}.txt")

    fichiers_cibles = []

    # 🔁 Parcours récursif des fichiers
    for root, _, files in os.walk(dossier_complet):
        for file in files:
            extension = os.path.splitext(file)[1].lower()
            if extension in extensions_cibles:
                fichiers_cibles.append(os.path.join(root, file))

    # ✏️ Écriture structurée dans le .txt
    with open(output_file, "w", encoding="utf-8") as sortie:
        for chemin_fichier in fichiers_cibles:
            try:
                with open(chemin_fichier, "r", encoding="utf-8", errors="ignore") as f:
                    contenu = f.read()
            except Exception as e:
                contenu = f"[Erreur lors de la lecture du fichier {chemin_fichier} : {e}]"

            chemin_relatif = os.path.relpath(chemin_fichier, dossier_complet)
            sortie.write(f"{chemin_relatif}\n\n")
            sortie.write(contenu)
            sortie.write("\n" + "_" * 120 + "\n\n")

print("\n✅ Compilation terminée. Tous les fichiers TXT sont générés.")