import  shutil
from    pathlib import Path

def clean_workspace():
    dirs_to_empty = [Path("jsons")]
    file_to_clear = Path("dependencies/codes.txt")
    for folder in dirs_to_empty:
        if folder.exists() and folder.is_dir():
            print(f"Cleaning directory: {folder}")
            for item in folder.iterdir():
                try:
                    if      item.is_file() or item.is_symlink() : item.unlink()
                    elif    item.is_dir()                       : shutil.rmtree(item)
                except Exception as e                           : print(f"[X] Failed to delete {item}: {e}")
        else: print(f"[X] Directory not found, skipping: {folder}")

    if file_to_clear.exists():
        print(f"[?] Clearing: {file_to_clear}")
        file_to_clear.write_text("")
    else: print(f"[X] File not found, skipping: {file_to_clear}")
    print("[✓] Cleanup complete")

if __name__ == "__main__": clean_workspace()