from pathlib import Path

def clean():
    tour_dir = Path(__file__).parent.absolute() / "tour"

    for folder in tour_dir.glob("*/*"):
        if folder.is_dir():
            print(f"[?] Cleaning subfolder: {folder}")
            
            for item in folder.rglob('*'):
                try:
                    if item.is_file() or item.is_symlink(): item.unlink()
                except Exception as e:  print(f"[X] Failed to delete {item}: {e}")

    for file_path in tour_dir.rglob('*.txt'):
        try:
            print(f"[?] Clearing contents of: {file_path}")
            file_path.write_text("")
        except Exception as e: print(f"[X] Failed to clear {file_path}: {e}")
        
    print("[✓] Cleanup complete")

if __name__ == "__main__": clean()