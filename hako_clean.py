from pathlib import Path

def delete_file(file: Path):
    try:
        if file.is_file() or file.is_symlink():
            print(f"[?] Deleting: {file}")
            file.unlink()
    except Exception as e: print(f"[X] Failed to delete {file}: {e}")

def clear_file(file: Path):
    try:
        print(f"[?] Clearing: {file}")
        file.write_text("")
    except Exception as e: print(f"[X] Failed to clear {file}: {e}")

def clean():
    script_dir  = Path(__file__).parent.absolute()
    json_dir    = script_dir / "jsons"
    tour_dir    = script_dir / "hako" / "tour"

    for folder in filter(Path.is_dir, script_dir.glob("hako-*")):
        for item in folder.rglob("*"): delete_file(item)

    for folder in filter(Path.is_dir, tour_dir.glob("*/*")):
        if folder.name == "txt": continue
        for item in folder.rglob("*"): delete_file(item)

    txt_files = set(script_dir.glob("*.txt")).union(tour_dir.rglob("*.txt"))
    for file_path in txt_files: clear_file(file_path)

    name_json = script_dir / "hako" / "template" / "name.json"
    if name_json.exists(): clear_file(name_json)

    dry_files = set(script_dir.glob("*.png")).union(json_dir.glob("*.json"))
    for file_path in dry_files: delete_file(file_path)

    print("[✓] Cleanup complete")

if __name__ == "__main__": clean()