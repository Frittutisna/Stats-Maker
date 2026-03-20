import  json
import  numpy           as      np
import  os
import  pandas          as      pd
import  re
import  tkinter         as      tk
from    collections     import  defaultdict, Counter
from    datetime        import  datetime
from    html2image      import  Html2Image
from    PIL             import  Image, ImageChops, ImageOps
from    tkinter         import  messagebox, ttk

EXCLUDED_TAGS = {
    "Female Protagonist", 
    "Male Protagonist", 
    "Primarily Female Cast", 
    "Primarily Male Cast", 
    "School", 
    "Heterosexual", 
    "Primarily Teen Cast",
    "Ensemble Cast"
}

def extract_year(vintage_str):
    if not vintage_str: return None
    years = re.findall(r'\d{4}', str(vintage_str))
    if not years: return None
    year_val    = float(years[0])
    season_map  = {"winter": 0.00, "spring": 0.25, "summer": 0.50, "fall": 0.75}
    v_lower     = str(vintage_str).lower()
    decimal     = 0.0
    for season, val in season_map.items():
        if season in v_lower:
            decimal = val
            break
    return year_val + decimal

def format_year(val):
    if val is None: return "N/A"
    year = int(val)
    frac = val - year
    if      frac < 0.25 : season = "Winter"
    elif    frac < 0.50 : season = "Spring"
    elif    frac < 0.75 : season = "Summer"
    else                : season = "Fall"
    return f"{season} {year}"

def load_aliases(script_dir):
    """Loads aliases into dictionary"""
    alias_map   = {}
    alias_path  = os.path.join(script_dir, "dependencies", "aliases.txt")
    if os.path.exists(alias_path):
        with open(alias_path, "r", encoding = "utf-8") as f:
            for line in f:
                if "," in line:
                    existing, new   = [x.strip() for x in line.split(",", 1)]
                    alias_map[new]  = existing
    return alias_map

def save_alias(script_dir, existing_name, new_name):
    """Appends new alias(es)"""
    dep_dir     = os.path.join(script_dir, "dependencies")
    os.makedirs(dep_dir, exist_ok = True)
    alias_path  = os.path.join(dep_dir, "aliases.txt")
    with open(alias_path, "a", encoding = "utf-8") as f: f.write(f"{existing_name}, {new_name}\n")

def trim_whitespace(image_path):
    """Removes empty white space from right and bottom of image"""
    with Image.open(image_path) as img:
        img     = img.convert("RGB")
        bg      = Image.new(img.mode, img.size, "white")
        diff    = ImageChops.difference(img, bg)
        bbox    = diff.getbbox()
        if bbox: 
            img = img.crop(bbox)
            img = ImageOps.expand(img, border = 10, fill = "white")
            img.save(image_path)

class PlayerAdditionDialog(tk.Toplevel):
    def __init__(self, parent, current_members, known_pool):
        super().__init__(parent)
        self.title("Manual Player Selection")
        self.added_players  = []
        self.known_pool     = sorted(list(known_pool - current_members))
        
        main_frame = ttk.Frame(self, padding = 10)
        main_frame.pack(fill = tk.BOTH, expand = True)

        tk.Label(main_frame, text = f"Lobby Count: {len(current_members)}, expected 8", font = ("Arial", 10, "bold")).pack()
        
        curr_text = "Detected: " + ", ".join(sorted(list(current_members)))
        tk.Label(main_frame, text = curr_text, wraplength = 400, fg = "blue").pack(pady = 5)
        
        tk.Label(main_frame, text = "Select player to add:").pack(anchor = tk.W)
        
        self.listbox = tk.Listbox(main_frame, height = 12, selectmode = tk.MULTIPLE)
        for name in self.known_pool: self.listbox.insert(tk.END, name)
        self.listbox.pack(fill = tk.BOTH, expand = True, pady = 5)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady = 10)
        ttk.Button(btn_frame, text="Add",       command = self.add_selected)    .pack(side = tk.LEFT, padx = 5)
        ttk.Button(btn_frame, text="Finish",    command = self.destroy)         .pack(side = tk.LEFT, padx = 5)
        
        self.grab_set()
        self.wait_window()

    def add_selected(self):
        selections = self.listbox.curselection()
        if not selections: return
        
        for i in selections:
            name = self.listbox.get(i)
            if name not in self.added_players: 
                self.added_players.append(name)
        
        messagebox.showinfo("Added", f"Added {len(selections)} players")
        self.destroy()

class SubSelectionDialog(tk.Toplevel):
    def __init__(self, parent, missing_roster):
        super().__init__(parent)
        self.title("Substitute Resolution")
        self.result = None
        tk.Label(self, text = "Multiple roster members are missing; which player(s) were subbed?").pack(padx = 20, pady = 10)
        self.listbox = tk.Listbox(self, height = len(missing_roster))
        self.listbox.pack(padx = 20, pady = 5, fill = tk.X)
        for m in missing_roster: self.listbox.insert(tk.END, m)
        ttk.Button(self, text = "Confirm", command = self.on_confirm).pack(pady = 10)
        self.grab_set()
        self.wait_window()
        
    def on_confirm(self):
        sel = self.listbox.curselection()
        if sel: 
            self.result = self.listbox.get(sel[0])
            self.destroy()

class ManualMatchDialog(tk.Toplevel):
    def __init__(self, parent, unknown_name, available_pool):
        super().__init__(parent)
        self.title("Manual Match Required")
        self.result = None
        ttk.Label(self, text = f"Could not find match for: '{unknown_name}'", font = ("Arial", 10, "bold")).pack(pady = 10)
        self.listbox = tk.Listbox(self, height = 15)
        self.listbox.pack(padx = 10, fill = tk.BOTH)
        for name in sorted(available_pool): self.listbox.insert(tk.END, name)
        ttk.Button(self, text = "Match Selected", command = self.on_match).pack(pady = 10)
        self.grab_set()
        self.wait_window()

    def on_match(self):
        sel = self.listbox.curselection()
        if sel: 
            self.result = self.listbox.get(sel[0])
            self.destroy()

def save_as_html_table(rows, md_path, title):
    html =  f"## {title}\n\n"
    html += '<table>\n'
    for i, row in enumerate(rows):
        style   =   ' style="text-align: center;"'
        html    +=  f'  <tr{style}>\n'
        for j, cell in enumerate(row):
            is_header   =   (i == 0)
            tag         =   'th style="text-align: center;"' if is_header else "td"
            content     =   str(cell)
            if is_header            : content = content.replace(" ", "<br>")
            if j == 0 or is_header  : content = f"<b>{content}</b>"
            html        +=  f'    <{tag}>{content}</{tag}>\n'
        html += '  </tr>\n'
    html += '</table>\n\n'
    with open(md_path, "a", encoding = "utf-8") as f: f.write(html)

def get_browser():
    """Detects common browser paths for HTML2Image"""
    paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" ,
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"       ,
        r"C:\Program Files\Google\Chrome\Application\chrome.exe"        ,
    ]
    for path in paths:
        if os.path.exists(path): return path
    return None

def export_df_to_png(df, path, filename, title):
    """Renders DataFrame as PNG"""
    browser_path = get_browser()
    if not browser_path:
        messagebox.showerror("[!] Error: Could not find Edge nor Chrome")
        return

    ascending_metrics   = ["Sevens", "Overs", "Average Overs"]
    descending_metrics  = [
        "Elo", 
        "Guess Rate", 
        "Solos", 
        "Doubles", 
        "Rigs", 
        "Points", 
        "Blocks", 
        "Rig Rate", 
        "OP GR", 
        "IN GR", 
        "ED GR", 
        "Rig GR", 
        "Off GR", 
        "Average GR", 
        "Rig Synergy", 
        "Off Synergy", 
        "Shared Rigs",
        "Total Solos",
    ]

    stats = {}
    for col in df.columns:
        if col in descending_metrics or col in ascending_metrics:
            numeric_col = pd.to_numeric(df[col].astype(str).str.replace('%', ''), errors = 'coerce').dropna()
            if not numeric_col.empty:
                max_val         = numeric_col.max()
                min_val         = numeric_col.min()
                counts          = numeric_col.value_counts()
                stats[col]      = {
                    'max'       : max_val, 
                    'min'       : min_val,
                    'show_max'  : counts.get(max_val, 0) <= 3,
                    'show_min'  : counts.get(min_val, 0) <= 3
                }

    rows_html = ""
    rows_html += '  <thead><tr>\n'
    for col in df.columns:
        content     =   str(col).replace(" ", "<br>")
        rows_html   +=  f'    <th>{content}</th>\n'
    rows_html += '  </tr></thead>\n<tbody>\n'
    
    for _, row in df.iterrows():
        rows_html += '  <tr>\n'
        for i, (col_name, cell) in enumerate(row.items()):
            content = str(cell)
            style   = ""
            if col_name in stats:
                val = pd.to_numeric(str(cell).replace('%', ''), errors = 'coerce')
                if pd.notnull(val):
                    is_max = (val == stats[col_name]['max']) and stats[col_name]['show_max']
                    is_min = (val == stats[col_name]['min']) and stats[col_name]['show_min']
                    if      col_name in descending_metrics:
                        if      is_max: style = ' style="color: #0056B3; font-weight: bold;"'
                        elif    is_min: style = ' style="color: #D95400; font-weight: bold;"'
                    elif    col_name in ascending_metrics:
                        if      is_max: style = ' style="color: #D95400; font-weight: bold;"'
                        elif    is_min: style = ' style="color: #0056B3; font-weight: bold;"'

            if i == 0: content = f"<b>{content}</b>"
            rows_html += f'    <td{style}>{content}</td>\n'
        rows_html += '  </tr>\n'
    rows_html += '</tbody>'

    width   = max(2000, len(df.columns) * 120)
    height  = max(2000, len(df)         * 60)

    hti = Html2Image(
        size                = (width, height), 
        browser_executable  = browser_path, 
        output_path         = path, 
        custom_flags        = ['--log-level=3', '--silent']
    )

    full_html = f"""
    <html>
    <head>
    <style>
        body {{ font-family : 'Segoe UI', Arial, sans-serif; background : white; display : inline-block; margin : 0; }}
        h2      {{margin: 10px 0 10px 5px; font-size: 30px; text-align: center;}}
        table   {{margin-left: 10px; border-collapse: collapse; width: auto;}}
        th      {{font-weight: bold; font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}}
        td      {{font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}}
    </style>
    </head>
    <body>
        <h2>{title}</h2>
        <table>{rows_html}</table>
    </body>
    </html>
    """

    hti.screenshot(html_str = full_html, save_as = filename)
    full_path = os.path.join(path, filename)
    try                     : trim_whitespace(full_path)
    except Exception as e   : print(f"[!] Error: Failed to trim {filename}: {e}")

def process_files():
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    json_dir    = os.path.join(script_dir, "jsons")
    json_paths  = []

    while True:
        if os.path.exists(json_dir) and os.path.isdir(json_dir): 
            json_paths = [os.path.join(json_dir, f) for f in os.listdir(json_dir) if f.endswith(".json")]
        if json_paths: break
        else:
            retry = messagebox.askyesno("Missing Files", "jsons folder not found or empty, click Yes to re-run")
            if not retry: return

    all_known_players = set()
    for path in json_paths:
        try:
            with open(path, encoding = "utf-8") as f:
                data = json.load(f)
                for s in data.get("songs", []):
                    for p   in s.get("correctGuessPlayers", []): all_known_players.add(p)
                    for ls  in s.get("listStates",          []): all_known_players.add(ls["name"])
        except: continue

    raw_assignments = {}
    team_rosters    = defaultdict(set)
    t1_lookup       = {}
    use_teams       = False
    codes_path      = os.path.join(script_dir, "dependencies", "codes.txt")
    
    if os.path.exists(codes_path):
        alias_map       = load_aliases(script_dir)
        player_elo_map  = {}
        with open(codes_path, "r", encoding = "utf-8") as f:
            lines = [l.strip() for l in f.readlines() if "(" in l]
            if lines:
                use_teams       = True
                available       = list(all_known_players)

                for t_idx, line in enumerate(lines, 1):
                    team_prefix_match   = re.match(r'^(?:\\s*)?([^:\[\d\(]+)\s*\([\d.]+\):', line)
                    explicit_team_name  = team_prefix_match.group(1).strip() if team_prefix_match else None

                    if ":" in line  : player_section = line.split(":", 1)[1]
                    else            : player_section = line

                    members = re.findall(r'([^\s(]+)\s*\((\d+\.\d+)\)', player_section)
                    if not members: continue

                    for i, (p_in, elo_val) in enumerate(members[:4]):
                        tier    = str(i + 1)
                        match   = next((n for n in available if n.lower() == p_in.lower()), None)

                        if not match and p_in in alias_map:
                            target_name = alias_map[p_in]
                            match       = next((n for n in available if n == target_name), None)
                        
                        if not match:
                            match = ManualMatchDialog(None, p_in, available).result
                            if match:
                                save_alias(script_dir, match, p_in)
                                alias_map[p_in] = match

                        if match:
                            raw_assignments [match] = (t_idx, tier)
                            player_elo_map  [match] = elo_val
                            team_rosters[t_idx].add(match)
                            if      match in available  : available.remove(match)
                            if      explicit_team_name  : t1_lookup[t_idx] = explicit_team_name
                            elif    tier == "1"         : t1_lookup[t_idx] = match

    song_participation          = defaultdict(int)
    correct_counts              = defaultdict(int)
    erigs_counts                = defaultdict(int)
    player_reverse_erigs        = defaultdict(int)
    player_two_eighths          = defaultdict(int)
    player_points               = defaultdict(int)
    player_blocks               = defaultdict(int)
    player_type_correct         = defaultdict(lambda: defaultdict(int)) 
    player_type_seen            = defaultdict(lambda: defaultdict(int))
    player_rigs                 = defaultdict(int)
    player_rigs_hit             = defaultdict(int)
    all_song_vintages           = []
    all_song_difficulties       = []
    total_correct_answers_sum   = 0
    total_blanks                = 0
    total_erigs                 = 0
    total_doubles               = 0
    total_sevens                = 0
    total_fulls                 = 0
    total_sweeps                = 0
    genre_counter               = Counter()
    tag_counter                 = Counter()
    player_list_vintages        = defaultdict(list)
    player_list_correct_counts  = defaultdict(list) 
    player_missed_erigs         = defaultdict(int)
    player_list_solos           = defaultdict(int)
    watched_only_valid          = True
    missing_list_count          = 0
    team_vintage                = defaultdict(list)
    team_correct_per_song       = defaultdict(list)
    team_onlist_synergy         = defaultdict(list)
    team_offlist_synergy        = defaultdict(list)
    team_shared_rig_pct         = defaultdict(list)
    team_solos                  = defaultdict(int)
    team_sweeps                 = defaultdict(int)
    team_overs                  = defaultdict(list)

    for path in json_paths:
        with open(path, encoding = "utf-8") as f: data = json.load(f)
        songs = data.get("songs", [])
        if not songs: continue
        
        raw_file_players    = {p for s in songs for p in s.get("correctGuessPlayers", [])} | {ls["name"] for s in songs for ls in s.get("listStates", [])}
        final_file_members  = set(raw_file_players)
        
        if use_teams:
            teams_in_file = {raw_assignments[p][0] for p in raw_file_players if p in raw_assignments}
            for t_id in teams_in_file:
                roster  = team_rosters[t_id]
                missing = [p for p in roster if p not in raw_file_players]
                if len([p for p in roster if p in raw_file_players]) == 3 and missing:
                    if len(missing) == 1: final_file_members.add(missing[0])
                    else:
                        res = SubSelectionDialog(None, missing).result
                        if res: final_file_members.add(res)
            
            if len(final_file_members) < 8:
                for t_id in teams_in_file:
                    for teammate in team_rosters[t_id]: final_file_members.add(teammate)

        while len(final_file_members) < 8:
            dialog = PlayerAdditionDialog(None, final_file_members, all_known_players)
            if not dialog.added_players:
                if messagebox.askyesno("Warning", "Still under 8 players, continue anyway?")    : break
                else                                                                            : continue
            for p in dialog.added_players                                                       : final_file_members.add(p)

        apply_rev           = (len(final_file_members) % 2 == 0)
        max_songs           = max(s.get("songNumber", 0) for s in songs)
        type_totals_file    = defaultdict(int)

        for song in songs:
            si = song.get("songInfo", {})
            st = si.get("type")

            if st in [1, 2, 3]                          : type_totals_file[st] += 1
            if isinstance(si.get("animeGenre"), list)   : genre_counter .update(si.get("animeGenre"))
            if isinstance(si.get("animeTags"),  list)   : tag_counter   .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])

            correct                     =   set(song.get("correctGuessPlayers", []))
            ls                          =   song.get("listStates", [])
            total_correct_answers_sum   +=  len(correct)
            year                        =   extract_year(si.get("vintage"))
            diff                        =   si.get("animeDifficulty")

            if isinstance(diff, (int, float))   : all_song_difficulties .append(diff)
            if year is not None                 : all_song_vintages     .append(year)
            if not ls                           : missing_list_count += 1

            song_riggers    = {p["name"] for p in ls}
            unique_lister   = ls[0]["name"] if len(ls) == 1 else None
            if unique_lister:
                player_list_solos[unique_lister] += 1
                is_solo = (len(correct) == 1 and list(correct)[0] == unique_lister)
                if not is_solo: player_missed_erigs[unique_lister] += 1

            if use_teams:
                teams_in_file_list = list({raw_assignments[p][0] for p in raw_file_players if p in raw_assignments})
                if len(teams_in_file_list) == 2:
                    tA, tB  = teams_in_file_list[0], teams_in_file_list[1]                    
                    c_teamA = correct & team_rosters[tA]
                    c_teamB = correct & team_rosters[tB]

                    if len(c_teamA) == len(team_rosters[tA]) and len(c_teamB) == 0:
                        team_sweeps[tA] += 1
                        total_sweeps    += 1
                    
                    if len(c_teamB) == len(team_rosters[tB]) and len(c_teamA) == 0:
                        team_sweeps[tB] += 1
                        total_sweeps    += 1

                    for cur_t, opp_t in [(tA, tB), (tB, tA)]:
                        cur_c, opp_c = correct & team_rosters[cur_t], correct & team_rosters[opp_t]
                        if not opp_c: 
                            for p in cur_c: player_points[p] += 1
                        if len(cur_c) == 1 and len(opp_c) > 0: player_blocks[list(cur_c)[0]] += 1

                for t_id in teams_in_file_list:
                    roster = team_rosters               [t_id]
                    c_on_t = correct        & roster
                    t_rigs = song_riggers   & roster
                    team_correct_per_song               [t_id].append(len(c_on_t)       / 4.0)
                    if year is not None: team_vintage   [t_id].append(year)
                    if t_rigs:
                        team_onlist_synergy             [t_id].append(len(c_on_t)       / 4.0)
                        team_shared_rig_pct             [t_id].append((len(t_rigs) - 1) / 3.0)
                        team_overs                      [t_id].append((len(correct), len(t_rigs)))
                    else: team_offlist_synergy          [t_id].append(len(c_on_t)       / 4.0)

            if len(final_file_members - correct) == 0: total_fulls += 1
            elif apply_rev and len(final_file_members - correct) == 1: 
                total_sevens += 1
                player_reverse_erigs[list(final_file_members - correct)[0]] += 1
            elif len(correct) == 2:
                total_doubles += 1
                for p in correct: player_two_eighths[p] += 1
            elif len(correct) == 1: 
                total_erigs += 1
                solo_winner = list(correct)[0]
                erigs_counts[solo_winner] += 1
                if solo_winner in raw_assignments: team_solos[raw_assignments[solo_winner][0]] += 1
            elif len(correct) == 0: total_blanks += 1

            for name in final_file_members:
                if name in correct:
                    correct_counts[name]                                += 1
                    if st in [1, 2, 3]: player_type_correct[name][st]   += 1
            if ls:
                for p in ls:
                    n                                                   =   p["name"]
                    player_rigs[n]                                      +=  1
                    if n in correct         : player_rigs_hit[n]             +=  1
                    if year is not None     : player_list_vintages[n].append(year)
                    player_list_correct_counts[n].append(len(correct))
        watched_only_valid = missing_list_count <= 5

        for name in final_file_members:
            song_participation[name]                        += max_songs
            for t in [1, 2, 3]: player_type_seen[name][t]   += type_totals_file[t]

    p_rows = []
    for name in song_participation:
        total           = song_participation    [name]
        correct         = correct_counts        [name]
        type_map        = {1: "OP GR", 3: "IN GR", 2: "ED GR"}        
        found_types     = {t_id: label for t_id, label in type_map.items() if any(player_type_seen[p][t_id] > 0 for p in song_participation)}
        active_types    = found_types if len(found_types) > 1 else {}

        row = {
            "Player"        : name, 
            "Guess Rate"    : correct / total if total else 0, 
            "Solos"         : erigs_counts          [name], 
            "Doubles"       : player_two_eighths    [name], 
            "Sevens"        : player_reverse_erigs  [name], 
            "Points"        : player_points         [name], 
            "Blocks"        : player_blocks         [name]
        }

        for t_id, label in active_types.items():
            seen        = player_type_seen[name][t_id]
            row[label]  = player_type_correct[name][t_id] / seen if seen else np.nan

        if watched_only_valid: row.update({
            "Rigs"          : player_rigs[name], 
            "Rig Rate"      : player_rigs[name] / total                                         if total                            else np.nan,
            "Rig GR"     : player_rigs_hit[name] / player_rigs[name]                         if player_rigs[name]                else np.nan,
            "Off GR"    : (correct - player_rigs_hit[name]) / (total - player_rigs[name])   if (total - player_rigs[name])      else np.nan,
            "Overs"         : np.mean(player_list_correct_counts[name])                         if player_list_correct_counts[name] else np.nan
        })

        if use_teams:
            elo = player_elo_map.get(name, "N/A")
            row = {"Elo": elo, **row} 
        p_rows.append(row)

    df_ps = pd.DataFrame(p_rows)
    if "Elo" in df_ps.columns:
        df_ps["Elo"]    = pd.to_numeric(df_ps["Elo"], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        cols            = list(df_ps.columns)
        cols.remove("Elo")
        idx         = cols.index("Guess Rate")
        cols.insert(idx, "Elo")
        df_ps           = df_ps[cols]
    df_ps = df_ps.sort_values("Guess Rate", ascending = False)

    df_display  = df_ps.copy()
    type_cols   = ["OP GR", "IN GR", "ED GR"]
    pct_cols    = ["Guess Rate"] + [c for c in type_cols if c in df_display.columns]

    if watched_only_valid: 
        pct_cols.extend(["Rig Rate", "Rig GR", "Off GR"])
        if "Overs" in df_display.columns:
            df_display["Overs"] = pd.to_numeric(df_display["Overs"], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

    for c in pct_cols: 
        df_display[c] = pd.to_numeric(df_display[c], errors = 'coerce').mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

    timestamp   = datetime.now().strftime("%y%m%d%H")
    png_dir     = os.path.join(script_dir, "archive", "png",    timestamp)
    md_path     = os.path.join(script_dir, "archive", "md",     f"{timestamp}.md")
    os.makedirs(png_dir,                    exist_ok = True)
    os.makedirs(os.path.dirname(md_path),   exist_ok = True)
    open(md_path, 'w').close()

    player_header   = [list(df_display.columns)]
    player_data     = df_display.values.tolist()
    plist           = list(song_participation.keys())
    max_solos_val   = max(erigs_counts.values())            if erigs_counts         else 0
    max_doubles_val = max(player_two_eighths.values())      if player_two_eighths   else 0
    max_sevens_val  = max(player_reverse_erigs.values())    if player_reverse_erigs else 0

    tied_solos      = [n for n, v in erigs_counts.items()           if v == max_solos_val   and v > 0]
    tied_doubles    = [n for n, v in player_two_eighths.items()     if v == max_doubles_val and v > 0]
    tied_sevens     = [n for n, v in player_reverse_erigs.items()   if v == max_sevens_val  and v > 0]

    def format_most_stat(names, value):
        if not names            : return "N/A"
        elif len(names) == 2    : return f"{names[0]} and {names[1]} ({value})"
        elif len(names) >= 3    : return f"{names[0]} and others ({value})"
        else                    : return f"{names[0]} ({value})"

    str_solos   = format_most_stat(tied_solos,      max_solos_val)
    str_doubles = format_most_stat(tied_doubles,    max_doubles_val)
    str_sevens  = format_most_stat(tied_sevens,     max_sevens_val)

    no_erig_pool        = [n for n in plist if erigs_counts         [n] ==  0 and song_participation[n] > 0]
    solos_pool          = [n for n in plist if erigs_counts         [n] >   0 and song_participation[n] > 0]
    list_players_pool   = [n for n in plist if player_list_solos    [n] >   0 and song_participation[n] > 0]

    best_no_erig        = sorted(no_erig_pool,  key = lambda x: (correct_counts[x] / song_participation[x]), reverse = True)    [0] if no_erig_pool else "N/A"
    worst_with_solos    = sorted(solos_pool,    key = lambda x: (correct_counts[x] / song_participation[x]), reverse = False)   [0] if solos_pool   else "N/A"

    if list_players_pool:
        min_missed_val              = min(player_missed_erigs[n] for n in list_players_pool)
        potential_least_missed      = [n for n in list_players_pool if player_missed_erigs[n] == min_missed_val]
        if len(potential_least_missed) > 1:
            max_list_solos_in_pool  = max(player_list_solos[n] for n in potential_least_missed)
            final_tied_least        = [n for n in potential_least_missed if player_list_solos[n] == max_list_solos_in_pool]
            winner_names            = format_most_stat(final_tied_least, None).split(" (")[0]
            list_solos_val          = max_list_solos_in_pool
        else:
            winner_names            = potential_least_missed[0]
            list_solos_val          = player_list_solos[winner_names]
        str_least_missed            = f"{winner_names} ({min_missed_val}/{list_solos_val})"

        max_missed_val              = max(player_missed_erigs[n] for n in list_players_pool)
        potential_most_missed       = [n for n in list_players_pool if player_missed_erigs[n] == max_missed_val]
        if len(potential_most_missed) > 1:
            min_list_solos_in_pool  = min(player_list_solos[n] for n in potential_most_missed)
            final_tied_most         = [n for n in potential_most_missed if player_list_solos[n] == min_list_solos_in_pool]
            winner_names_most       = format_most_stat(final_tied_most, None).split(" (")[0]
            list_solos_val_most     = min_list_solos_in_pool
        else:
            winner_names_most       = potential_most_missed[0]
            list_solos_val_most     = player_list_solos[winner_names_most]
        str_missed                  = f"{winner_names_most} ({max_missed_val}/{list_solos_val_most})"
    else: 
        str_least_missed            = "N/A"
        str_missed                  = "N/A"

    tour_stats = [
        ["Average Vintage",     format_year(round(np.mean(all_song_vintages), 2))],
        ["Average Difficulty",  f"{np.mean(all_song_difficulties):.2f}"],
        ["Average GR",          f"{100 * (total_correct_answers_sum / sum(song_participation.values())):.2f}"],
        ["Total Blanks",        total_blanks],
        ["Total Solos",         total_erigs],
        ["Total Doubles",       total_doubles],
        ["Total Sevens",        total_sevens],
        ["Total Fulls",         total_fulls],
        ["Total Sweeps",        total_sweeps],
        ["Most Popular Genre",  f"{genre_counter    .most_common(1)[0][0]} ({genre_counter  .most_common(1)[0][1]})" if genre_counter   else "N/A"],
        ["Most Popular Tag",    f"{tag_counter      .most_common(1)[0][0]} ({tag_counter    .most_common(1)[0][1]})" if tag_counter     else "N/A"],
        ["Most Solos",          str_solos],
        ["Most Doubles",        str_doubles],
        ["Most Sevens",         str_sevens]
    ]

    if best_no_erig     != "N/A": tour_stats.append(["Highest GR without Solos",    f"{best_no_erig} ({100      * (correct_counts[best_no_erig]     / song_participation[best_no_erig])     :.2f})"])
    if worst_with_solos != "N/A": tour_stats.append(["Lowest GR with Solos",        f"{worst_with_solos} ({100  * (correct_counts[worst_with_solos] / song_participation[worst_with_solos]) :.2f}, {erigs_counts[worst_with_solos]})"])

    if watched_only_valid: 
        tour_stats.append(["Most Missed Solos",     str_missed])
        tour_stats.append(["Least Missed Solos",    str_least_missed])
    
    extra_content       = [["Statistic", "Value"]] + tour_stats
    team_stats_content  = []
    tier_stats_content  = []

    if use_teams and watched_only_valid:
        team_headers = [
            "Team", 
            "Average Vintage", 
            "Average GR", 
            "Rig Synergy", 
            "Off Synergy", 
            "Shared Rigs", 
            "Total Solos", 
            "Team Sweeps",
            "Average Overs"
        ]
        team_stats_content  = [team_headers]
        stats_list          = []
        for t_id in team_correct_per_song.keys():
            if team_overs[t_id]:
                total_rig_overs     = sum(over * rigs for over, rigs in team_overs[t_id])
                total_rigs          = sum(rigs for _, rigs in team_overs[t_id])
                weighted_overs      = total_rig_overs / total_rigs if total_rigs > 0 else 0.0
            else: weighted_overs    = 0.0
            stats_list.append({
                "id"        : t_id,
                "vintage"   : np.mean(team_vintage          [t_id]) if team_vintage             [t_id] else np.nan,
                "avg"       : np.mean(team_correct_per_song [t_id]) if team_correct_per_song    [t_id] else 0.0,
                "onlist"    : np.mean(team_onlist_synergy   [t_id]) if team_onlist_synergy      [t_id] else 0.0,
                "offlist"   : np.mean(team_offlist_synergy  [t_id]) if team_offlist_synergy     [t_id] else 0.0,
                "shared"    : np.mean(team_shared_rig_pct   [t_id]) if team_shared_rig_pct      [t_id] else 0.0,
                "solos"     : team_solos                    [t_id],
                "sweeps"    : team_sweeps                   [t_id],
                "overs"     : weighted_overs
            })

        stats_list.sort(key = lambda x: x["avg"], reverse = True)
        for item in stats_list:
            row = [
                t1_lookup.get(item["id"], f"Team {item['id']}"), 
                format_year(item["vintage"]), 
                f"{item ['avg']     * 100   :.2f}", 
                f"{item ['onlist']  * 100   :.2f}", 
                f"{item ['offlist'] * 100   :.2f}", 
                f"{item ['shared']  * 100   :.2f}", 
                item    ['solos'], 
                item    ['sweeps'],
                f"{item ['overs']           :.2f}"
            ]
            team_stats_content.append(row)

    if use_teams:
        tier_stats_content  = [["Tier", "Attacker", "Blocker"]]
        tiers               = sorted({v[1] for v in raw_assignments.values() if v[1] != "N/A"})
        for tr in tiers:
            tier_players    = [n for n, assign in raw_assignments.items() if assign[1] == tr]
            tdf             = df_ps[df_ps["Player"].isin(tier_players)].copy()
            
            if not tdf.empty:
                max_pts     = tdf["Points"].max()
                max_blk     = tdf["Blocks"].max()
                top_atks    = tdf[tdf["Points"] == max_pts]
                top_blks    = tdf[tdf["Blocks"] == max_blk]
                atk_row     = top_atks.sort_values("Guess Rate",    ascending = False).iloc[0]
                blk_row     = top_blks.sort_values("Blocks",        ascending = False).iloc[0]
                
                atk_val     = atk_row['Guess Rate'] if isinstance(atk_row['Guess Rate'], float) else 0.0
                blk_val     = blk_row['Guess Rate'] if isinstance(blk_row['Guess Rate'], float) else 0.0

                if len(top_atks) > 1    : atk_display = f"{atk_row['Player']} ({atk_row['Points']}, {atk_val * 100:.2f}%)"
                else                    : atk_display = f"{atk_row['Player']} ({atk_row['Points']})"
                if len(top_blks) > 1    : blk_display = f"{blk_row['Player']} ({blk_row['Blocks']}, {blk_val * 100:.2f}%)"
                else                    : blk_display = f"{blk_row['Player']} ({blk_row['Blocks']})"
                tier_stats_content.append([tr, atk_display, blk_display])

    watched_content = []
    if watched_only_valid:
        e_list = sorted([(n, np.mean(player_list_correct_counts [n])) for n in plist if player_list_correct_counts  [n]], key = lambda x: x[1], reverse = True) [:3]
        h_list = sorted([(n, np.mean(player_list_correct_counts [n])) for n in plist if player_list_correct_counts  [n]], key = lambda x: x[1])                 [:3]
        z_list = sorted([(n, np.mean(player_list_vintages       [n])) for n in plist if player_list_vintages        [n]], key = lambda x: x[1], reverse = True) [:3]
        b_list = sorted([(n, np.mean(player_list_vintages       [n])) for n in plist if player_list_vintages        [n]], key = lambda x: x[1])                 [:3]
        watched_content.append(["Rank", "Easiest", "Hardest", "Newest", "Oldest"])
        for i in range(3):
            row = [f"{i + 1}"]
            row.append(f"{e_list[i][0]} ({e_list                [i][1]:.2f})"   if i < len(e_list) else "N/A")
            row.append(f"{h_list[i][0]} ({h_list                [i][1]:.2f})"   if i < len(h_list) else "N/A")
            row.append(f"{z_list[i][0]} ({format_year(z_list    [i][1])})"      if i < len(z_list) else "N/A")
            row.append(f"{b_list[i][0]} ({format_year(b_list    [i][1])})"      if i < len(b_list) else "N/A")
            watched_content.append(row)

    save_as_html_table(player_header + player_data, md_path, "Player Statistics")
    export_df_to_png(df_display, png_dir, "Player.png", "Player Statistics")

    if team_stats_content:
        save_as_html_table(team_stats_content, md_path, "Team Statistics")
        export_df_to_png(
            pd.DataFrame(team_stats_content[1:], columns = team_stats_content[0]), 
            png_dir, "Team.png", "Team Statistics"
        )

    if tier_stats_content:
        save_as_html_table(tier_stats_content, md_path, "Tier Bests")
        export_df_to_png(
            pd.DataFrame(tier_stats_content[1:], columns = tier_stats_content[0]), 
            png_dir, "Tier.png", "Tier Bests"
        )

    if extra_content:
        save_as_html_table(extra_content, md_path, "Tour Statistics")
        export_df_to_png(
            pd.DataFrame(extra_content[1:], columns = extra_content[0]), 
            png_dir, "Tour.png", "Tour Statistics"
        )

    if watched_content: 
        save_as_html_table(watched_content, md_path, "Watched Statistics")
        export_df_to_png(
            pd.DataFrame(watched_content[1:], columns = watched_content[0]), 
            png_dir, "Watched.png", "Watched Statistics"
        )

    if messagebox.askyesno("Success", f"Saved as Markdown and PNGs, click Yes to delete JSONs"):
        for path in json_paths:
            try     : os.remove(path)
            except  : pass

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    process_files()