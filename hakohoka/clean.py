from pathlib import Path

def clean_workspace():
    base_dir        = Path(__file__).parent.absolute().parent
    indices         = range(3)
    dirs_to_empty   = [base_dir / f"tours/{i}/jsons"     for i in indices] + [base_dir / f"tours/{i}/arpia" for i in indices]   + [base_dir / f"tours/{i}/hakohoka" for i in indices]
    files_to_clear  = [base_dir / f"tours/{i}/codes.txt" for i in indices] + [base_dir / "tours/aliases.txt"]                   + [base_dir / "tours/notes.txt"]

    for folder in dirs_to_empty:
        if not folder.is_dir():
            print(f"[X] Directory not found, skipping: {folder}")
            continue

        print(f"[?] Cleaning: {folder}")

        for item in folder.rglob('*'):
            try:
                if item.is_file() or item.is_symlink(): item.unlink()

            except Exception as e: print(f"[X] Failed to delete {item}: {e}")

    for file_path in files_to_clear:
        if file_path.is_file():
            print(f"[?] Clearing contents of: {file_path}")
            file_path.write_text("")

        else: print(f"[X] File not found, skipping: {file_path}")
        
    print("[✓] Cleanup complete")

if __name__ == "__main__": clean_workspace()