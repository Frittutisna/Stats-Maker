from pathlib import Path

def clean():
    script_dir  = Path(__file__).parent.absolute()
    tour_dir    = script_dir / "tour"

    for folder in tour_dir.glob("*/*"):
        if folder.is_dir():
            print(f"[?] Cleaning subfolder: {folder}")

            for item in folder.rglob('*'):
                try:
                    if item.is_file() or item.is_symlink(): item.unlink()

                except Exception as e:  print(f"[X] Failed to delete {item}: {e}")

    all_txt_files = set(script_dir.glob('*.txt')).union(tour_dir.rglob('*.txt'))

    for file_path in all_txt_files:
        try:
            print(f"[?] Clearing contents of: {file_path}")
            file_path.write_text("")

        except Exception as e: print(f"[X] Failed to clear {file_path}: {e}")

    for file_path in script_dir.glob('*.png'):
        try:
            print(f"[?] Deleting PNG: {file_path}")
            file_path.unlink()

        except Exception as e: print(f"[X] Failed to delete {file_path}: {e}")

    print("[✓] Cleanup complete")

if __name__ == "__main__": clean()