import os
import subprocess

# Chemin vers ILSpyCmd (modifie selon ton installation)
ilspy_path = r"C:\Users\Nicolas JF Martin\.dotnet\tools\ilspycmd.exe"

# Dossier contenant les DLL à analyser
dll_folder = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\leo_salvador\07-OX_TOOLS\OxTools\Helpers\ExcelConnector\Assemblies"

# Dossier de sortie principal
output_dir = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation_2"

# Creation du dossier "transposition_dll" s'il n'existe pas
transposition_dir = os.path.join(output_dir, "transposition_dll")
os.makedirs(transposition_dir, exist_ok=True)

# Fichier log pour les DLL en echec
log_file_path = os.path.join(output_dir, "log.txt")

# Verification de la presence de ILSpyCmd
if not os.path.exists(ilspy_path):
    print("Erreur: ILSpyCmd n'est pas trouve.")
    exit()

# Recuperer la liste des DLL dans le dossier
dll_files = [f for f in os.listdir(dll_folder) if f.endswith(".dll")]

# Verifier s'il y a des DLL a traiter
if not dll_files:
    print("Erreur: Aucune DLL trouvee dans le dossier.")
    exit()

# Ouvrir le fichier log
with open(log_file_path, "w", encoding="utf-8") as log_file:
    log_file.write("DLL en echec de decompilation:\n\n")

    # Boucle sur chaque DLL pour la decompiler
    for idx, dll_file in enumerate(dll_files, start=1):
        dll_path = os.path.join(dll_folder, dll_file)
        dll_name = os.path.splitext(dll_file)[0]

        # Creation d’un sous-dossier pour chaque DLL
        dll_output_dir = os.path.join(transposition_dir, dll_name)
        os.makedirs(dll_output_dir, exist_ok=True)

        # Commande pour generer les fichiers .cs
        cmd = [ilspy_path, dll_path, "-p", "-o", dll_output_dir]  # "-p" pour projet, "-o" pour dossier de sortie

        # Afficher l'avancement
        print(f"[INFO] ({idx}/{len(dll_files)}) Traitement de {dll_file}...")

        # Executer la commande
        result = subprocess.run(cmd, capture_output=True, text=True)

        # Verifier le resultat
        if result.returncode == 0:
            print(f"[OK] {dll_file} traite avec succes. Resultats dans: {dll_output_dir}")
        else:
            print(f"[ERREUR] Echec de la decompilation de {dll_file}.")
            log_file.write(f"{dll_file}\n")  # Ajouter au fichier log

print(f"[COMPLET] Tous les fichiers DLL ont ete traites. Verifie {log_file_path} pour les erreurs.")