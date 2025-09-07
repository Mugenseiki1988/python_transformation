import os
import glob

# =========================
# CONFIG
# =========================
ROOT_PATHS = [
    r"C:\Program Files (x86)\AVEVA",
    r"D:\E3D.2.1",
]

# Fichier de sortie unique
OUTPUT_FILE = r"D:\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\FORMATION E3D ADMIN\ETUDE AVEVA UIC ETC\UIC XML TRANSFORMATION\3-xlsx_tcd_output\concat_xmls_output.txt"

# Tri des chemins pour reproductibilité (True/False)
SORT_RESULTS = True

# Encodages à essayer à la lecture (dans l'ordre)
ENCODINGS_TRY = ["utf-8", "utf-16", "utf-16-le", "utf-16-be", "cp1252", "latin-1"]


def read_text_keep_indentation(file_path):
    """
    Lit le fichier texte en essayant plusieurs encodages,
    en conservant strictement l'indentation / le contenu brut.
    """
    last_err = None
    for enc in ENCODINGS_TRY:
        try:
            with open(file_path, "r", encoding=enc, errors="strict") as f:
                return f.read()
        except Exception as e:
            last_err = e
            continue
    # Dernier recours : lecture binaire puis décodage permissif
    try:
        with open(file_path, "rb") as fb:
            raw = fb.read()
        return raw.decode("utf-8", errors="replace")
    except Exception:
        # Si vraiment impossible, on renvoie une trace minimale
        return f"[ERREUR LECTURE: {file_path}]\n{str(last_err)}"


def gather_xml_files(root_paths):
    """
    Récupère la liste dédupliquée de tous les .xml sous les chemins fournis.
    """
    files = []
    for rp in root_paths:
        files.extend(glob.glob(os.path.join(rp, "**", "*.xml"), recursive=True))
    # déduplication
    files = list(set(files))
    if SORT_RESULTS:
        files.sort(key=lambda p: (p.lower(), len(p)))
    return files


def build_concatenation(xml_paths):
    """
    Construit la chaîne finale avec la mise en forme demandée.
    """
    parts = []
    for p in xml_paths:
        try:
            filename = os.path.basename(p)
            content = read_text_keep_indentation(p)
            block = f"{filename}\n\n{p}\n\n{content}\n\n------"
            parts.append(block)
        except Exception as e:
            # On loggue l'erreur sous forme de bloc distinct pour garder la trace
            err_block = f"[ERREUR TRAITEMENT]\n\n{p}\n\n{str(e)}\n\n------"
            parts.append(err_block)
    return "\n".join(parts)


def main():
    xml_files = gather_xml_files(ROOT_PATHS)
    if not xml_files:
        print("Aucun fichier XML trouvé dans les répertoires fournis.")
        return

    print(f"{len(xml_files)} fichiers XML détectés. Génération du fichier de sortie...")
    final_text = build_concatenation(xml_files)

    # Ecriture du fichier de sortie en UTF-8 (BOM non nécessaire)
    # On n’altère pas l’indentation des blocs (déjà conservée à la lecture).
    with open(OUTPUT_FILE, "w", encoding="utf-8", newline="\n") as out:
        out.write(final_text)

    print(f"Fichier généré : {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
