import shutil, subprocess, sys
from pathlib import Path

def main():
    script_dir          = Path(__file__).parent.absolute()
    tour_dir            = script_dir        / "tour"
    public_repo_dir     = script_dir.parent / "Stats-Public"
    target_jsons_dir    = public_repo_dir   / "jsons"
    target_codes_file   = public_repo_dir   / "codes.txt"
    target_script       = public_repo_dir   / "public.py"

    available_tours = ["0", "1", "2"]
    print("[?] Which tour should be processed for final results?")
    for idx, tour in enumerate(available_tours): print(f"[{idx}] Tour {tour}")
        
    while True:
        try:
            choice = input("Select a number (0-2): ").strip()

            if choice in ["0", "1", "2"]:
                selected_tour = available_tours[int(choice)]
                break

            print("[X] Invalid selection, please choose 0, 1, or 2")

        except (KeyboardInterrupt, EOFError):
            print("[X] Operation cancelled")
            sys.exit(0)

    selected_tour_path  = tour_dir              / selected_tour
    source_jsons_dir    = selected_tour_path    / "json"
    source_codes_file   = selected_tour_path    / "code.txt"

    print(f"[?] Processing Tour {selected_tour}")

    if target_jsons_dir.exists(): shutil.rmtree(target_jsons_dir)
    target_jsons_dir.mkdir(parents = True, exist_ok = True)

    if source_jsons_dir.exists() and any(source_jsons_dir.glob("*.json")):
        for json_file in source_jsons_dir.glob("*.json"): shutil.copy(json_file, target_jsons_dir / json_file.name)
        print("[✓] Copied JSONs")

    else: print("[!] JSONs not found")

    if source_codes_file.exists():
        shutil.copy(source_codes_file, target_codes_file)
        print("[✓] Copied code.txt")

    else: print("[!] code.txt not found")

    if not target_script.exists():
        print("[X] public.py not found")
        sys.exit(1)

    print("[?] Running public.py")
    
    try:
        subprocess.run([sys.executable, str(target_script)], cwd = str(public_repo_dir), check = True)
        print(f"[✓] Public data generated")

    except subprocess.CalledProcessError as e: print(f"[X] public.py exited with an error status: {e}")

if __name__ == "__main__": main()