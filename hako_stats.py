import gspread, ijson, json, os, shutil, sys, urllib.request
import tkinter as tk

from hako.analyzer       import TourAnalyzer
from hako.help.config    import DIR_CREDS, DIR_TOURS, FILE_CHANT
from hako.help.dialog    import TourSelectionDialog
from pathlib             import Path

def sync_chanting(tour_dir_path):
    cred_file   = os.path.join("hako", "help", DIR_CREDS, "credentials.json")
    auth_file   = os.path.join("hako", "help", DIR_CREDS, "authorized_user.json")
    sheet_name  = "NGM Stats Export v2"

    try:
        gc              = gspread.oauth(credentials_filename = cred_file, authorized_user_filename = auth_file)
        sheet           = gc.open(sheet_name)
        rows            = sheet.worksheet("MiscData").get_all_values()
        chanting_ids    = set()

        for row in rows[1:]:
            if not row: continue
            value = str(row[0]).strip()
            if value: chanting_ids.add(value)

        tour_dir_path.mkdir(parents = True, exist_ok = True)
        chant_file_path = tour_dir_path / FILE_CHANT
        with open(chant_file_path, "w", encoding = "utf-8") as f: f.write("\n".join(sorted(list(chanting_ids), key = lambda x: int(x) if x.isdigit() else x)))

    except Exception as e: print(f"Failed to download chanting song IDs: {e}")

def extract_unique_names(template_dir):
    name_file_path = template_dir / "Name.json"

    if not name_file_path.exists() or os.path.getsize(name_file_path) == 0:
        print("[?] Name.json is missing or empty, extracting from libraryMasterList")

        url             = "https://animemusicquiz.com/libraryMasterList"
        unique_names    = set()
        headers         = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        request         = urllib.request.Request(url, headers=headers)

        try:
            with urllib.request.urlopen(request) as response:
                parser = ijson.kvitems(response, "animeMap")

                for _, anime_data in parser:
                    main_names = anime_data.get("mainNames", {})
                    if not main_names: continue

                    ja_name = main_names.get("JA")
                    en_name = main_names.get("EN")

                    if ja_name: unique_names.add(ja_name.strip())
                    if en_name: unique_names.add(en_name.strip())

            flat_list = sorted(list(unique_names))
            with open(name_file_path, "w", encoding = "utf-8") as out_f: json.dump(flat_list, out_f, ensure_ascii = False, indent = 4)
            print(f"[✓] Success! Extracted {len(flat_list)} unique names to: {name_file_path}")
            
        except Exception as e: print(f"[X] Failed to extract unique names: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    script_directory    = Path(__file__).parent.absolute()
    tour_folder_path    = script_directory / "hako" / DIR_TOURS
    chant_txt_file      = tour_folder_path / FILE_CHANT    
    has_valid_tour      = False
    template_path       = script_directory / "hako" / "help" / "template"

    extract_unique_names(template_path)

    for json_dir in tour_folder_path.glob("*/json"):
        if any(json_dir.glob("*.json")):
            has_valid_tour = True
            break

    if not has_valid_tour:
        root_jsons = script_directory / "jsons"

        if not root_jsons.exists() or not any(root_jsons.glob("*.json")):
            print("[X] Error: No JSON files found in hako/tour/*/json or root/jsons")
            sys.exit(1)

        print("[?] No valid tours found, initializing Tour 0 with root/jsons")

        target_json_dir = tour_folder_path / "0" / "json"
        target_json_dir.mkdir(parents = True, exist_ok = True)

        for j_file in root_jsons.glob("*.json"): shutil.copy(j_file, target_json_dir / j_file.name)

        root_codes      = script_directory / "codes.txt"
        target_codes    = tour_folder_path / "0" / "code.txt"

        if root_codes.exists():
            target_codes.parent.mkdir(parents = True, exist_ok = True)
            shutil.copy(root_codes, target_codes)

    selection_dialog    = TourSelectionDialog(root, ["0", "1", "2"])
    selected_tours      = selection_dialog.selected_tours

    if not chant_txt_file.exists() or os.path.getsize(chant_txt_file) == 0: sync_chanting(tour_folder_path)
    
    if selected_tours:
        analyzer_pool = []
        
        for tour_id in selected_tours:
            analyzer = TourAnalyzer(tour_id)
            is_valid = analyzer.prepare_configuration()

            if is_valid: analyzer_pool.append(analyzer)

            else:
                print(f"Tour {tour_id} failed configuration checks, halting pipeline execution")
                root.destroy()
                sys.exit(1)
                
        for analyzer in analyzer_pool: analyzer.process_and_generate()