import os
import re
import fnmatch
import pandas as pd
from tqdm import tqdm

# === Paramètres ===
search_directories = [
    r"C:\\Program Files (x86)\\AVEVA\\Everything3D2.10",
    r"D:\\E3D.2.1\\MEIUI",
    r"D:\\E3D.2.1\\MEILIB"
]

file_extensions = ["*.pmlfrm", "*.pmlobj", "*.pmlcmd", "*.pmlfnc", "*.pmlmac", "*"]
exact_terms = ["import", "namespace", "member", "object"]
partial_terms = ["!dll"]

scan_specific_file = True
specific_file_path = r"C:\\Program Files (x86)\\AVEVA\\Everything3D2.10\\OXTOOLS\\Helpers\\ExcelConnector\\Forms\\excelConnectorInterface.pmlfrm"

output_file = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\DLL_dans_fichiers_pml_1.xlsx"

excel_input = r"D:\\BUREAU-BUREAU-BUREAU-BUREAU-BUREAU\\FORMATION E3D ADMIN\\ETUDE AVEVA UIC ETC\\code_unique_dll_static_ILSpy.xlsx"


def load_dll_scan(excel_path):
    df = pd.read_excel(excel_path)
    df = df[["Concatener 2"]].dropna()
    df.columns = ["dll_path"]
    return df


def scan_file_for_terms(file_path):
    results = []
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
        for line in lines:
            line_lower = line.lower()
            if any(term in line_lower for term in exact_terms + partial_terms):
                results.append(line.strip())
    except Exception:
        pass
    return results


def get_files(directory):
    all_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            ext = os.path.splitext(file)[1]
            if ext:
                if any(fnmatch.fnmatch(file.lower(), pattern.lower()) for pattern in file_extensions if pattern):
                    all_files.append(file_path)
            elif "" in file_extensions:
                all_files.append(file_path)
    return all_files


def determine_extension(file):
    base = os.path.basename(file)
    ext = os.path.splitext(base)[1]
    return ext[1:] if ext else ""


def validate_dll_segment(dll_segment, df_pml_scan):
    for line in df_pml_scan["line"]:
        if "import" in line.lower():
            if re.search(rf"import\s+[\"']{re.escape(dll_segment)}[\"']", line, re.IGNORECASE):
                print(f"\t→ DLL directe validée par import : {dll_segment}")
                return True
            if re.search(rf"\$!dllPath", line, re.IGNORECASE):
                for subline in df_pml_scan["line"]:
                    if re.search(rf"!dllName\s*=\s*\|?{re.escape(dll_segment)}\|?", subline, re.IGNORECASE):
                        print(f"\t→ DLL dynamique validée via !dllName = {dll_segment}")
                        return True
    return False


def validate_class_segment(class_segment, df_pml_scan):
    class_escaped = re.escape(class_segment)
    pattern_member = rf"\bmember\b.*\b{class_escaped}\b"
    pattern_object = rf"\bobject\b.*\b{class_escaped}\b\s*\("

    for line in df_pml_scan["line"]:
        if re.search(pattern_member, line, re.IGNORECASE):
            print(f"\t→ Classe validée via 'member' : {class_segment}")
            return True
        if re.search(pattern_object, line, re.IGNORECASE):
            print(f"\t→ Classe validée via 'object' : {class_segment}")
            return True
    return False


def validate_namespace_segment(namespace_segment, df_pml_scan):
    for line in df_pml_scan["line"]:
        if re.search(rf"\bnamespace\b.*{re.escape(namespace_segment)}", line, re.IGNORECASE):
            print(f"\t→ Namespace validé : {namespace_segment}")
            return True
    return False


def validate_and_extract_segments(dll_path, df_pml_scan):
    match = re.match(r"^(.*?)/(.*?)/(.+\.cs)/(.*)$", dll_path)
    if not match:
        return None

    dll, namespace, class_cs, visibility = match.groups()
    cls = class_cs.replace(".cs", "")

    if not visibility.strip():
        return None
    visibility_cleaned = visibility.split()[0].lower()

    if not validate_dll_segment(dll, df_pml_scan):
        return None
    if not validate_class_segment(cls, df_pml_scan):
        return None
    if not validate_namespace_segment(namespace, df_pml_scan):
        return None

    return {
        "DLL": dll,
        "Namespace": namespace,
        "Class": cls,
        "Visibility": visibility_cleaned,
        "Variable": f"{dll}/{namespace}/{cls}.cs/{visibility}"
    }


def detect_imports(df_pml_scan):
    imports = set()
    for line in df_pml_scan["line"]:
        match_direct = re.search(r"import\s+[\"'](.+?)[\"']", line, re.IGNORECASE)
        if match_direct:
            imports.add(match_direct.group(1))
    return list(imports)


# === MAIN ===
df_dll_scan = load_dll_scan(excel_input)
df_collect = pd.DataFrame(columns=[
    "Fichier", "Nom", "Dossier", "Extension",
    "DLL", "Namespace", "Class", "Visibility", "Variable"
])
match_count = 0

if scan_specific_file:
    print(f"\n▶ Scan du fichier : {specific_file_path}")
    file = specific_file_path
    lines = scan_file_for_terms(file)
    df_pml_scan = pd.DataFrame(lines, columns=["line"])
    print(f" ▶ Lignes scannées : {len(df_pml_scan)}")
    print("📦 DLL importées dans le fichier :", detect_imports(df_pml_scan))

    for _, row in df_dll_scan.iterrows():
        dll_path = str(row["dll_path"])
        match_info = validate_and_extract_segments(dll_path, df_pml_scan)
        if match_info:
            if match_info["Variable"] in df_collect["Variable"].values:
                continue
            match_count += 1
            df_collect = pd.concat([df_collect, pd.DataFrame([{
                "Fichier": file,
                "Nom": os.path.basename(file),
                "Dossier": os.path.dirname(file),
                "Extension": determine_extension(file),
                "DLL": match_info["DLL"],
                "Namespace": match_info["Namespace"],
                "Class": match_info["Class"],
                "Visibility": match_info["Visibility"],
                "Variable": match_info["Variable"]
            }])], ignore_index=True)

print(f"\n▶ Total DLL matchés : {match_count}")
if not df_collect.empty:
    df_collect.to_excel(output_file, index=False)
    print(f"📁 Export terminé → {output_file}")
else:
    print("⚠️ Aucun match détecté. Fichier Excel vide.")