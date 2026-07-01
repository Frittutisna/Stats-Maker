from pathlib import Path

def clean():
    script_dir  = Path(__file__).parent.absolute()
    json_dir    = script_dir / "jsons"
    tour_dir    = script_dir / "hako" / "tour"

    for folder in tour_dir.glob("*/*"):
        if folder.is_dir():
            print(f"[?] Cleaning subfolder: {folder}")

            for item in folder.rglob('*'):
                try:
                    if item.is_file() or item.is_symlink(): item.unlink()

                except Exception as e:  print(f"[X] Failed to delete {item}: {e}")

    txt_files = set(script_dir.glob('*.txt')).union(tour_dir.rglob('*.txt'))

    for file_path in txt_files:
        try:
            print(f"[?] Clearing contents of: {file_path}")
            file_path.write_text("")

        except Exception as e: print(f"[X] Failed to clear {file_path}: {e}")

    name_json = script_dir / "hako" / "help" / "template" / "Name.json"

    if name_json.exists():
        try:
            print(f"[?] Clearing contents of Name.json")
            name_json.write_text("")

        except Exception as e: print(f"[X] Failed to clear Name.json: {e}")

    dry_files = set(script_dir.glob('*.png')).union(json_dir.glob('*.json'))

    for file_path in dry_files:
        try:
            print(f"[?] Deleting Dry's files: {file_path}")
            file_path.unlink()

        except Exception as e: print(f"[X] Failed to delete {file_path}: {e}")

    print("[✓] Cleanup complete")

if __name__ == "__main__": clean()