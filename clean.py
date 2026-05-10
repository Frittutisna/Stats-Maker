from pathlib import Path

def clean_workspace():
    indices         = range(3)
    dirs_to_empty   = [Path(f"tours/{i}/jsons")     for i in indices] + [Path(f"tours/{i}/output") for i in indices]
    files_to_clear  = [Path(f"tours/{i}/codes.txt") for i in indices]

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