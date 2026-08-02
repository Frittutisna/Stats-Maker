import glob, json, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

all_json_files  = glob.glob("*.json")
files_to_stitch = sorted([f for f in all_json_files if not f.startswith("stitched-")])

if not files_to_stitch: print("No JSON files found to stitch")
else:
    combined_songs  = []
    room_name       = ""
    start_time      = ""

    for i, filename in enumerate(files_to_stitch):
        with open(filename, 'r', encoding = "utf-8") as f: data = json.load(f)
        
        if i == 0:
            room_name   = data.get('roomName',  'Unknown Room')
            start_time  = data.get('startTime', 'Unknown Time')

        for song in data.get('songs', []):
            new_song                = song.copy()
            new_song['songNumber']  = len(combined_songs) + 1

            combined_songs.append(new_song)

    stitched_data = {
        "roomName"  : room_name,
        "startTime" : start_time,
        "songs"     : combined_songs
    }

    output_name = f"stitched-{files_to_stitch[0]}"
    with open(output_name, 'w', encoding = "utf-8") as f: json.dump(stitched_data, f, indent = 4)

    print(f"Stitched {len(combined_songs)} songs from {len(files_to_stitch)} files into {output_name}")
    print("Cleaning up source files")

    for filename in files_to_stitch:
        try:
            os.remove(filename)
            print(f"Deleted {filename}")
        except Exception as e: print(f"Failed to delete {filename}: {e}")