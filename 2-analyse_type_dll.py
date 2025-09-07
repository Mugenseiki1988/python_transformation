import os
import pefile
import csv

# Répertoire contenant les DLLs
directory = r"D:\aveva_install\AVEVA_extensions_tri\DLL"  # Remplace avec ton répertoire
output_csv = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\DLL decompilation\dll_classification.csv"

# Liste toutes les DLLs dans le répertoire
dll_files = [f for f in os.listdir(directory) if f.endswith(".dll")]

# Création du fichier CSV et écriture de l'en-tête
with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(["Nom de la DLL", "Type"])

    # Analyse des DLLs et écriture des résultats dans le CSV
    for dll in dll_files:
        dll_path = os.path.join(directory, dll)
        try:
            pe = pefile.PE(dll_path)
            if hasattr(pe.OPTIONAL_HEADER, 'DATA_DIRECTORY') and pe.OPTIONAL_HEADER.DATA_DIRECTORY[14].VirtualAddress != 0:
                dll_type = "DLL .NET"
            else:
                dll_type = "DLL NATIVE (C/C++)"
            
            # Affichage et écriture dans le CSV
            print(f"{dll} -> {dll_type}")
            writer.writerow([dll, dll_type])
        
        except Exception as e:
            print(f"Erreur avec {dll}: {e}")
            writer.writerow([dll, f"Erreur: {e}"])

print(f"\n Analyse terminée. Résultats enregistrés dans : {output_csv}")