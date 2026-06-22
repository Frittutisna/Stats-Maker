import gspread, os, sys
import tkinter as tk

from analyzer       import TourAnalyzer
from help.config    import DIR_CREDS, DIR_TOURS, FILE_CHANT
from help.dialog    import TourSelectionDialog
from pathlib        import Path

def sync_chanting(tour_dir_path):
    cred_file   = os.path.join("help", DIR_CREDS, "credentials.json")
    auth_file   = os.path.join("help", DIR_CREDS, "authorized_user.json")
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

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    script_directory    = Path(__file__).parent.absolute()
    tour_folder_path    = script_directory / DIR_TOURS
    chant_txt_file      = tour_folder_path / FILE_CHANT    
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
                
        for analyzer in analyzer_pool: analyzer.process_and_generate() # Test