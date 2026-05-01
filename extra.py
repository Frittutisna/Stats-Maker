import  json
import  os
import  re
import  numpy       as      np
import  pandas      as      pd
import  tkinter     as      tk
from    collections import  defaultdict, Counter
from    datetime    import  datetime
from    html2image  import  Html2Image
from    PIL         import  Image, ImageChops, ImageOps
from    tkinter     import  messagebox, ttk

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
    if val is None or val == "N/A": return "N/A"
    year = int(val)
    frac = val - year
    if      frac < 0.25 : season = "Winter"
    elif    frac < 0.50 : season = "Spring"
    elif    frac < 0.75 : season = "Summer"
    else                : season = "Fall"
    return f"{season} {year}"

def trim_whitespace(image_path):
    with Image.open(image_path) as img:
        img     = img.convert("RGB")
        bg      = Image.new(img.mode, img.size, "white")
        diff    = ImageChops.difference(img, bg)
        bbox    = diff.getbbox()
        if bbox: 
            img = img.crop(bbox)
            img = ImageOps.expand(img, border = 10, fill = "white")
            img.save(image_path)

def get_browser():
    paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    ]
    for path in paths:
        if os.path.exists(path): return path
    return None

def load_aliases(script_dir):
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
    dep_dir = os.path.join(script_dir, "dependencies")
    os.makedirs(dep_dir, exist_ok = True)
    alias_path = os.path.join(dep_dir, "aliases.txt")
    with open(alias_path, "a", encoding = "utf-8") as f: f.write(f"{existing_name}, {new_name}\n")

def get_json_paths(script_dir):
    json_dir = os.path.join(script_dir, "jsons")
    while True:
        if os.path.exists(json_dir) and os.path.isdir(json_dir):
            paths = [os.path.join(json_dir, f) for f in os.listdir(json_dir) if f.endswith(".json")]
            if paths: return paths
        if not messagebox.askyesno("Missing Files", "jsons folder not found or empty, click Yes to re-run"): return []

def get_all_known_players(json_paths):
    all_players = set()
    appearances = defaultdict(set)
    for path in json_paths:
        try:
            with open(path, encoding = "utf-8") as f:
                data = json.load(f)
                for s in data.get("songs", []):
                    for p in s.get("correctGuessPlayers", []): 
                        all_players.add (p)
                        appearances     [p].add(path)
                    for ls in s.get("listStates", []): 
                        all_players.add (ls["name"])
                        appearances     [ls["name"]].add(path)
        except: continue
    return all_players, appearances

def load_team_data(script_dir, all_known_players):
    codes_path = os.path.join(script_dir, "dependencies", "codes.txt")
    if not os.path.exists(codes_path): return False, {}, {}, {}, defaultdict(set)
    alias_map       = load_aliases(script_dir)
    player_elo_map  = {}
    raw_assignments = {}
    team_rosters    = defaultdict(set)
    t1_lookup       = {}
    available       = list(all_known_players)
    with open(codes_path, "r", encoding="utf-8") as f: all_lines = f.readlines()

    for line in all_lines:
        found_elos = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', line)
        for p_in, elo_val in found_elos:
            match = next((n for n in all_known_players if n.lower() == p_in.lower()), None)
            if not match and p_in in alias_map: match = next((n for n in all_known_players if n == alias_map[p_in]), None)
            if not match and ("[" in line or "Subs:" in line):
                match = ManualMatchDialog(None, p_in, available).result
                if match:
                    save_alias(script_dir, match, p_in)
                    alias_map[p_in] = match
            if match: player_elo_map[match.lower()] = elo_val

    team_idx = 1
    for line in all_lines:
        if "[" not in line or "]" not in line: continue
        team_prefix_match   = re.match(r'^(?:\\s*)?([^:\[\d\(]+)\s*\([\d.-]+\):', line)
        explicit_name       = team_prefix_match.group(1).strip() if team_prefix_match else None
        player_section      = line.split(":", 1)[1] if ":" in line else line
        members             = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', player_section)

        for i, (p_in, _) in enumerate(members[:4]):
            tier    = str(i + 1)
            match   = next((n for n in all_known_players if n.lower() == p_in.lower() or (p_in in alias_map and n == alias_map[p_in])), None)
            if match:
                raw_assignments[match.lower()] = (team_idx, tier)
                team_rosters[team_idx].add(match)
                if      match in available  : available.remove(match)
                if      explicit_name       : t1_lookup[team_idx] = explicit_name
                elif    tier == "1"         : t1_lookup[team_idx] = match
        team_idx += 1

    return True, player_elo_map, raw_assignments, t1_lookup, team_rosters

class PlayerAdditionDialog(tk.Toplevel):
    def __init__(self, parent, current_members, known_pool):
        super().__init__(parent)
        self.title("Manual Player Selection")
        self.added_players  = []
        self.known_pool     = sorted(list(known_pool - current_members))
        main_frame          = ttk.Frame(self, padding = 10)
        main_frame.pack(fill = tk.BOTH, expand = True)
        tk.Label(main_frame, text = f"Lobby Count: {len(current_members)}, expected 8", font = ("Arial", 10, "bold")).pack()
        curr_text = "Detected: " + ", ".join(sorted(list(current_members)))
        tk.Label(main_frame, text = curr_text, wraplength = 400, fg = "blue")   .pack(pady      = 5)
        tk.Label(main_frame, text = "Select player to add:")                    .pack(anchor    = tk.W)
        self.listbox = tk.Listbox(main_frame, height = 10, selectmode = tk.MULTIPLE)
        for name in self.known_pool: self.listbox.insert(tk.END, name)
        self.listbox.pack(fill = tk.BOTH, expand = True, pady = 5)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady = 10)
        ttk.Button(btn_frame, text = "Add",     command = self.add_selected)    .pack(side = tk.LEFT, padx = 5)
        ttk.Button(btn_frame, text = "Finish",  command = self.destroy)         .pack(side = tk.LEFT, padx = 5)
        self.grab_set(); self.wait_window()
    def add_selected(self):
        selections = self.listbox.curselection()
        for i in selections:
            name = self.listbox.get(i)
            if name not in self.added_players: self.added_players.append(name)
        messagebox.showinfo("Added", f"Added {len(selections)} players"); self.destroy()

class SubSelectionDialog(tk.Toplevel):
    def __init__(self, parent, missing_roster):
        super().__init__(parent)
        self.title("Substitute Resolution")
        self.result = None
        tk.Label(self, text="Multiple roster members are missing; which player(s) were subbed?").pack(padx = 20, pady = 10)
        self.listbox = tk.Listbox(self, height = len(missing_roster))
        self.listbox.pack(padx = 20, pady = 5, fill=tk.X)
        for m in missing_roster: self.listbox.insert(tk.END, m)
        ttk.Button(self, text = "Confirm", command = self.on_confirm).pack(pady = 10)
        self.grab_set(); self.wait_window()
    def on_confirm(self):
        sel = self.listbox.curselection()
        if sel: self.result = self.listbox.get(sel[0]); self.destroy()

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
        self.grab_set(); self.wait_window()
    def on_match(self):
        sel = self.listbox.curselection()
        if sel: self.result = self.listbox.get(sel[0]); self.destroy()

def export_df_to_png(df, path, filename, title):
    browser_path = get_browser()
    if not browser_path: messagebox.showerror("[!] Error: Could not find Edge nor Chrome"); return

    ascending_metrics = [
        "Sevens",
        "Overs",
        "Average Overs"
    ]

    descending_metrics  = [
        "Elo",
        "Guess Rate",
        "Solos",
        "Doubles",
        "Rigs",
        "Rig Delta",
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
        "Total Solos"
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

    df          = df.reset_index(drop = True)
    border_rows = []

    if "Guess Rate" in df.columns:
        gr_vals     = pd.to_numeric(df["Guess Rate"].astype(str).str.replace('%', ''), errors = 'coerce').tolist()
        is_watched  = "Rigs" in df.columns
        thresholds  = [28.0, 18.0, 12.0, 6.0] if is_watched else [28.0, 19.0, 8.0]
        
        for thresh in thresholds:
            found_idx = -1

            for idx, val in enumerate(gr_vals):
                if pd.notnull(val) and val >= thresh: found_idx = idx
            
            if found_idx != -1 and found_idx < len(df) - 1: border_rows.append(found_idx)

    rows_html = "<thead><tr>" + "".join([f"<th>{str(col).replace(' ', '<br>')}</th>" for col in df.columns]) + "</tr></thead><tbody>"
    
    for idx, row in df.iterrows():
        border_style    =   "border-bottom: 3px solid black;" if idx in border_rows else ""
        rows_html       +=  "<tr>"

        for i, (col_name, cell) in enumerate(row.items()):
            content, style_parts = str(cell), []
            if border_style: style_parts.append(border_style)
            
            if col_name in stats:
                val = pd.to_numeric(str(cell).replace('%', ''), errors = 'coerce')
                if pd.notnull(val):
                    is_max = (val == stats[col_name]['max']) and stats[col_name]['show_max']
                    is_min = (val == stats[col_name]['min']) and stats[col_name]['show_min']

                    if      col_name in descending_metrics:
                        if      is_max: style_parts.append("color: #0056B3; font-weight: bold;")
                        elif    is_min: style_parts.append("color: #D95400; font-weight: bold;")

                    elif    col_name in ascending_metrics:
                        if      is_max: style_parts.append("color: #D95400; font-weight: bold;")
                        elif    is_min: style_parts.append("color: #0056B3; font-weight: bold;")
            
            style_attr          =   f' style="{" ".join(style_parts)}"' if style_parts else ""
            if i == 0: content  =   f"<b>{content}</b>"
            rows_html           +=  f"<td{style_attr}>{content}</td>"
        rows_html += "</tr>"
    rows_html += "</tbody>"

    hti = Html2Image(
        size                = (max(2000, len(df.columns) * 120), max(2000, len(df) * 60)),
        browser_executable  = browser_path,
        output_path         = path,
        custom_flags        = ['--log-level=3', '--silent']
    )

    full_html = f"<html><head><style>body {{font-family: 'Segoe UI', Arial, sans-serif; background: white; display: inline-block; margin: 0;}} h2 {{margin: 10px 0 10px 5px; font-size: 30px; text-align: center;}} table {{margin-left: 10px; border-collapse: collapse; width: auto;}} th {{font-weight: bold; font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}} td {{font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}}</style></head><body><h2>{title}</h2><table>{rows_html}</table></body></html>"
    hti.screenshot(html_str = full_html, save_as = filename)

    try                     : trim_whitespace(os.path.join(path, filename))
    except Exception as e   : print(f"[!] Error trimming {filename}: {e}")

def fuse_pngs(png_dir):
    files   = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "Watched": "Watched.png"}
    paths   = {k: os.path.join(png_dir, v)  for k, v in files.items() if os.path.exists(os.path.join(png_dir, v))}
    imgs    = {k: Image.open(v)             for k, v in paths.items()}
    if not imgs: return
    r_keys  = [k for k in ["Team", "Tier", "Watched"] if k in imgs]
    
    if len(r_keys) == 1:
        tw, th  = (imgs["Tour"].width, imgs["Tour"].height) if "Tour" in imgs else (0, 0)
        ok      = r_keys[0]
        ow      = imgs[ok].width
        oh      = imgs[ok].height
        fused   = Image.new("RGB", (max(tw, ow), th + (10 if th else 0) + oh), "white")
        if "Tour" in imgs: fused.paste(imgs["Tour"], (0, 0))
        fused.paste(imgs[ok], (0, th + 10 if th else 0))
    else:
        tw, th  = (imgs["Tour"].width, imgs["Tour"].height) if "Tour" in imgs else (0, 0)
        rw      = max([imgs[k].width        for k in r_keys])       if r_keys else 0
        rh      = sum([imgs[k].height + 10  for k in r_keys]) - 10  if r_keys else 0
        fused   = Image.new("RGB", (tw + (10 if tw and rw else 0) + rw, max(th, rh)), "white")
        if "Tour" in imgs: fused.paste(imgs["Tour"], (0, 0))
        cx, cy = (tw + 10 if tw else 0), 0
        for k in r_keys: fused.paste(imgs[k], (cx, cy)); cy += imgs[k].height + 10
    
    final_p = os.path.join(png_dir, "Extra.png")
    fused.save(final_p)

    try     : trim_whitespace(final_p)
    except  : pass

    for p in paths.values():
        try     : os.remove(p)
        except  : pass

def process_files():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_paths = get_json_paths(script_dir)
    if not json_paths: return

    all_known_players, player_json_appearances                          = get_all_known_players (json_paths)
    use_teams, player_elo_map, raw_assignments, t1_lookup, team_rosters = load_team_data        (script_dir, all_known_players)

    s_part      = defaultdict(int)
    c_counts    = defaultdict(int)
    e_counts    = defaultdict(int)
    p_rev_e     = defaultdict(int)
    p_two_e     = defaultdict(int)
    p_pts       = defaultdict(int)
    p_blks      = defaultdict(int)
    p_type_c    = defaultdict(lambda: defaultdict(int))
    p_type_s    = defaultdict(lambda: defaultdict(int))
    p_rigs      = defaultdict(int)
    p_rigs_h    = defaultdict(int)
    p_l_vint    = defaultdict(list)
    p_l_corr    = defaultdict(list)
    p_m_erigs   = defaultdict(int)
    p_l_solos   = defaultdict(int)
    t_vint      = defaultdict(list)
    t_c_ps      = defaultdict(list)
    t_on_syn    = defaultdict(list)
    t_off_syn   = defaultdict(list)
    t_sh_rig    = defaultdict(list)
    t_solos     = defaultdict(int)
    t_sweeps    = defaultdict(int)
    t_overs     = defaultdict(list)

    all_diff            = []
    all_vint            = []
    tot_c               = 0
    tot_blanks          = 0
    tot_erigs           = 0
    tot_doubles         = 0
    tot_sevens          = 0
    tot_fulls           = 0
    tot_sweeps          = 0
    genre_c             = Counter()
    tag_c               = Counter()
    missing_list_count  = 0
    found_types         = set()

    for path in json_paths:
        with open(path, encoding = "utf-8") as f: data = json.load(f)
        songs = data.get("songs", [])
        if not songs: continue

        raw_f_players = {p for s in songs for p in s.get("correctGuessPlayers", [])} | {ls["name"] for s in songs for ls in s.get("listStates", [])}
        final_members = set(raw_f_players)

        if use_teams:
            t_in_f = {raw_assignments[p.lower()][0] for p in raw_f_players if p.lower() in raw_assignments}
            for t_id in t_in_f:
                roster  = team_rosters[t_id]
                missing = [p for p in roster if p not in raw_f_players]
                if len([p for p in roster if p in raw_f_players]) == 3 and missing:
                    res = SubSelectionDialog(None, missing).result if len(missing) > 1 else missing[0]
                    if res: final_members.add(res)
            if len(final_members) < 8:
                for t_id in t_in_f: final_members.update(team_rosters[t_id])

        while len(final_members) < 8:
            dialog = PlayerAdditionDialog(None, final_members, all_known_players)
            if not dialog.added_players:
                if      messagebox.askyesno("Warning", "Still under 8 players, continue anyway?")   : break
                else                                                                                : continue
            final_members.update(dialog.added_players)

        apply_rev       = (len(final_members) % 2 == 0)
        max_s           = max(s.get("songNumber", 0) for s in songs)
        f_type_totals   = defaultdict(int)

        for song in songs:
            st = song.get("songInfo", {}).get("type")
            if st in [1, 2, 3]: 
                f_type_totals   [st] += 1
                found_types.add (st)

        players_in_this_file = {p for s in songs for p in s.get("correctGuessPlayers", [])} | {ls["name"] for s in songs for ls in s.get("listStates", [])}
        for name in final_members:
            if name in players_in_this_file:
                s_part[name] += max_s
                for t in [1, 2, 3]: p_type_s[name][t] += f_type_totals[t]

        for song in songs:
            si = song.get("songInfo", {}); st = si.get("type")
            if isinstance(si.get("animeGenre"), list)   : genre_c   .update(si.get("animeGenre"))
            if isinstance(si.get("animeTags"),  list)   : tag_c     .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])
            
            correct =   set(song.get    ("correctGuessPlayers", []))
            ls      =   song.get        ("listStates",          [])
            tot_c   +=  len(correct)
            yr      =   extract_year(si.get("vintage"))
            dfc     =   si.get("animeDifficulty")
            if isinstance(dfc, (int, float))    : all_diff.append(dfc)
            if yr is not None                   : all_vint.append(yr)
            if not ls                           : missing_list_count += 1

            s_riggers = {p["name"] for p in ls}
            if len(ls) == 1:
                u_lister = ls[0]["name"]
                p_l_solos[u_lister] += 1
                if not (len(correct) == 1 and list(correct)[0] == u_lister): p_m_erigs[u_lister] += 1

            if use_teams:
                t_in_f_list = list({raw_assignments[p.lower()][0] for p in raw_f_players if p.lower() in raw_assignments})
                if len(t_in_f_list) == 2:
                    tA, tB = t_in_f_list[0], t_in_f_list[1]
                    cA, cB = correct & team_rosters[tA], correct & team_rosters[tB]
                    if len(cA) == 4 and not cB: t_sweeps[tA] += 1; tot_sweeps += 1
                    if len(cB) == 4 and not cA: t_sweeps[tB] += 1; tot_sweeps += 1
                    for cur, opp in [(tA, tB), (tB, tA)]:
                        cC, oC = correct & team_rosters[cur], correct & team_rosters[opp]
                        if not oC: 
                            for p in cC: p_pts[p] += 1
                        if len(cC) == 1 and len(oC) > 0: p_blks[list(cC)[0]] += 1
                
                for t_id in t_in_f_list:
                    ros     = team_rosters[t_id]
                    c_on_t  = correct & ros
                    t_rg    = s_riggers & ros
                    t_c_ps[t_id].append(len(c_on_t) / 4.0)
                    if yr is not None: t_vint[t_id].append(yr)
                    if t_rg:
                        t_on_syn    [t_id].append(len(c_on_t)       / 4.0)
                        t_sh_rig    [t_id].append((len(t_rg) - 1)   / 3.0)
                        t_overs     [t_id].append((len(correct), len(t_rg)))
                    else: t_off_syn [t_id].append(len(c_on_t)       / 4.0)

            if      len(final_members - correct) == 0               : tot_fulls     += 1
            elif    apply_rev and len(final_members - correct) == 1 : tot_sevens    += 1; p_rev_e[list(final_members - correct)[0]] += 1
            elif    len(correct) == 2                               : tot_doubles   += 1; [p_two_e.__setitem__(p, p_two_e[p] + 1) for p in correct]
            elif    len(correct) == 1:
                tot_erigs       +=  1
                sw              =   list(correct)[0]
                e_counts[sw]    +=  1
                if sw.lower() in raw_assignments: t_solos[raw_assignments[sw.lower()][0]] += 1
            elif len(correct) == 0: tot_blanks += 1

            for name in final_members:
                if name in correct:
                    c_counts[name] += 1
                    if st in [1, 2, 3]: p_type_c[name][st] += 1
            if ls:
                for p in ls:
                    n = p["name"]; p_rigs[n] += 1
                    if n in correct     : p_rigs_h[n] += 1
                    if yr is not None   : p_l_vint[n].append(yr)
                    p_l_corr[n].append(len(correct))

    watched_valid   = missing_list_count <= 5
    med_jsons       = int(np.median([len(v) for v in player_json_appearances.values()])) if player_json_appearances else 0

    if      med_jsons == 3                                                                  : stage = "Mid-Tour"
    elif    (len(s_part) <= 20 and med_jsons >= 6) or (len(s_part) > 20 and med_jsons >= 5) : stage = "Final"
    else                                                                                    : stage = f"R{int(round(med_jsons))}"

    prefix          = ""
    type_map        = {1: "OP", 3: "IN", 2: "ED"}
    found_abbrs     = sorted([type_map[t] for t in found_types if t in type_map])
    all_types_found = set(type_map.keys()).issubset(found_types)
    
    if watched_valid:
        if all_types_found: prefix  = "Watched Tour, "
        else:
            type_str                = "-".join(found_abbrs) if found_abbrs else ""
            prefix                  = f"Watched {type_str} Tour, "
    else:
        if all_types_found: prefix  = "Random Tour, "
        else:
            type_str                = "/".join(found_abbrs) if found_abbrs else ""
            prefix                  = f"Random {type_str} Tour, "

    timestamp   = datetime.now().strftime("%y%m%d%H")
    png_dir     = os.path.join(script_dir, "archive", timestamp)
    os.makedirs(png_dir, exist_ok = True)

    create_player_report(
        s_part,
        c_counts,
        e_counts,
        p_two_e,
        p_rev_e,
        p_pts,
        p_blks,
        p_type_c,
        p_type_s,
        p_rigs,
        p_rigs_h,
        p_l_corr,
        use_teams,
        player_elo_map,
        watched_valid,
        stage,
        png_dir,
        player_json_appearances,
        prefix
    )

    create_tour_report(
        all_vint,
        all_diff,
        tot_c,
        s_part,
        tot_blanks,
        tot_erigs,
        tot_doubles,
        tot_sevens,
        tot_fulls,
        tot_sweeps,
        use_teams,
        genre_c,
        tag_c,
        e_counts,
        c_counts,
        p_two_e,
        p_rev_e,
        p_l_solos,
        p_m_erigs,
        watched_valid,
        stage,
        png_dir
    )
    
    if use_teams and watched_valid:
        create_team_report(t_c_ps, t_vint, t_on_syn, t_off_syn, t_sh_rig, t_solos, t_overs, t1_lookup, png_dir)
        create_tier_report(s_part, raw_assignments, c_counts, p_pts, p_blks, png_dir)
    
    if watched_valid: create_watched_report(s_part, p_l_corr, p_l_vint, png_dir)

    fuse_pngs(png_dir)
    messagebox.showinfo("Success", f"Saved PNGs to archive/{timestamp}")

def create_player_report(
        s_part,
        c_counts,
        e_counts,
        p_two_e,
        p_rev_e,
        p_pts,
        p_blks,
        p_type_c,
        p_type_s,
        p_rigs,
        p_rigs_h,
        p_l_corr,
        use_teams,
        elo_map,
        watched_valid,
        stage,
        png_dir,
        player_json_appearances,
        prefix
    ):
    p_rows      = []
    type_labels = {1: "OP GR", 3: "IN GR", 2: "ED GR"}
    
    active_types_list   = [t for t in type_labels if any(p_type_s[p][t] > 0 for p in s_part)]
    active_types        = {t: type_labels[t] for t in active_types_list} if len(active_types_list) > 1 else {}

    appear_counts   = [len(v) for v in player_json_appearances.values()]
    med_appear      = int(np.median(appear_counts)) if appear_counts else 0

    for name in s_part:
        total           = s_part    [name]
        correct         = c_counts  [name]
        actual_jsons    = len(player_json_appearances.get(name, []))
        display_name    = f"{name} ◀▶" if actual_jsons < med_appear else name
        row             = {"Player": display_name}
        if use_teams: row["Elo"] = elo_map.get(name.lower(), "N/A")

        row.update({
            "Guess Rate"    : correct / total if total else 0,
            "Solos"         : e_counts  [name],
            "Doubles"       : p_two_e   [name],
            "Sevens"        : p_rev_e   [name]
        })

        if use_teams: row.update({"Points": p_pts[name], "Blocks": p_blks[name]})

        for t_id, label in active_types.items():
            seen        = p_type_s[name][t_id]
            row[label]  = p_type_c[name][t_id] / seen if seen else np.nan

        if watched_valid: 
            row.update({
                "Rigs"      : p_rigs                [name],
                "Rig Rate"  : p_rigs                [name]  / total                     if total                    else np.nan,
                "Rig Delta" : (correct - p_rigs     [name]) / correct                   if correct                  else np.nan,
                "Rig GR"    : p_rigs_h              [name]  / p_rigs[name]              if p_rigs           [name]  else np.nan,
                "Off GR"    : (correct - p_rigs_h   [name]) / (total - p_rigs[name])    if (total - p_rigs  [name]) else np.nan,
                "Overs"     : np.mean(p_l_corr[name])                                   if p_l_corr         [name]  else np.nan
            })

        p_rows.append(row)

    df                                      = pd.DataFrame(p_rows).sort_values("Guess Rate", ascending = False)
    if "Elo" in df.columns  : df["Elo"]     = pd.to_numeric(df["Elo"],      errors = 'coerce')          .map(lambda x: f"{x:.2f}"           if pd.notnull(x) else "N/A")
    pct_cols                                = ["Guess Rate"] + list(active_types.values()) + (["Rig Rate", "Rig Delta", "Rig GR", "Off GR"] if watched_valid else [])
    for c in pct_cols       : df[c]         = pd.to_numeric(df[c],          errors = 'coerce').mul(100) .map(lambda x: f"{x:.2f}"           if pd.notnull(x) else "N/A")
    if watched_valid        : df["Overs"]   = pd.to_numeric(df["Overs"],    errors = 'coerce')          .map(lambda x: f"{x:.2f}"           if pd.notnull(x) else "N/A")
    
    export_df_to_png(df, png_dir, "Player.png", f"{prefix}Player Statistics, {stage}")

def create_tour_report(
        all_vint,
        all_diff,
        tot_c,
        s_part,
        tot_blanks,
        tot_erigs,
        tot_doubles,
        tot_sevens,
        tot_fulls,
        tot_sweeps,
        use_teams,
        genre_c,
        tag_c,
        e_counts,
        c_counts,
        p_two_e,
        p_rev_e,
        p_l_solos,
        p_m_erigs,
        watched_valid,
        stage,
        png_dir
    ):
    
    def format_most(names, val):
        if not names: return "N/A"
        winner  = sorted(names, key = lambda x: (c_counts[x] / s_part[x])   if s_part[x]        else 0)[0]
        gr      = (c_counts[winner] / s_part[winner]) * 100                 if s_part[winner]   else 0
        return f"{winner} ({val}{f', {gr:.2f}'                              if len(names) > 1   else ''})"

    tour_stats = [
        ["Median Vintage",      format_year(round(np.median(all_vint), 2))      if all_vint else "N/A"],
        ["Average Difficulty",  f"{np.mean(all_diff):.2f}"                      if all_diff else "N/A"],
        ["Average GR",          f"{100 * (tot_c / sum(s_part.values())):.2f}"   if s_part   else "0.00"],
        ["Total Blanks",        tot_blanks], 
        ["Total Solos",         tot_erigs], 
        ["Total Doubles",       tot_doubles], 
        ["Total Sevens",        tot_sevens], 
        ["Total Fulls",         tot_fulls]
    ]
    
    if use_teams: tour_stats.append(["Total Sweeps", tot_sweeps])
        
    tour_stats.extend([
        ["Most Popular Genre",  f"{genre_c  .most_common(1)[0][0]} ({genre_c    .most_common(1)[0][1]})" if genre_c else "N/A"],
        ["Most Popular Tag",    f"{tag_c    .most_common(1)[0][0]} ({tag_c      .most_common(1)[0][1]})" if tag_c   else "N/A"],
        ["Most Solos",          format_most([n for n, v in e_counts .items() if v == max(e_counts   .values(), default = 0) and v > 0], max(e_counts    .values(), default = 0))],
        ["Most Doubles",        format_most([n for n, v in p_two_e  .items() if v == max(p_two_e    .values(), default = 0) and v > 0], max(p_two_e     .values(), default = 0))],
        ["Most Sevens",         format_most([n for n, v in p_rev_e  .items() if v == max(p_rev_e    .values(), default = 0) and v > 0], max(p_rev_e     .values(), default = 0))]
    ])

    plist   = list(s_part.keys())
    no_erig = sorted([n for n in plist if e_counts[n] == 0 and s_part[n] > 0], key = lambda x: c_counts[x] / s_part[x], reverse = True)
    if no_erig: tour_stats.append(["Highest GR without Solos", f"{no_erig[0]} ({100 * (c_counts[no_erig[0]] / s_part[no_erig[0]]):.2f})"])
    
    solos_pool = sorted([n for n in plist if e_counts[n] > 0 and s_part[n] > 0], key = lambda x: c_counts[x] / s_part[x])
    if solos_pool:
        w_solo = solos_pool[0]
        tour_stats.append(["Lowest GR with Solos", f"{w_solo} ({100 * (c_counts[w_solo] / s_part[w_solo]):.2f}, {e_counts[w_solo]})"])

    if watched_valid:
        conv = []
        for n in [p for p in plist if p_l_solos[p] > 0]:
            h = p_l_solos[n] - p_m_erigs[n]
            conv.append({'n': n, 'p': 100 * h / p_l_solos[n], 'h': h, 't': p_l_solos[n]})
        if conv:
            b = sorted(conv, key = lambda x: (x['p'], x['t']),  reverse = True)     [0]
            w = sorted(conv, key = lambda x: (x['p'], -x['t']), reverse = False)    [0]
            tour_stats.append(["Best Solo Rig Converter",   f"{b['n']} ({b['p']:.2f}%, {b['h']}/{b['t']})"])
            tour_stats.append(["Worst Solo Rig Converter",  f"{w['n']} ({w['p']:.2f}%, {w['h']}/{w['t']})"])

    export_df_to_png(pd.DataFrame(tour_stats, columns = ["Statistic", "Value"]), png_dir, "Tour.png", "Tour Statistics")

def create_team_report(t_c_ps, t_vint, t_on_syn, t_off_syn, t_sh_rig, t_solos, t_overs, t1_lookup, png_dir):
    stats_list = []
    for t_id in t_c_ps.keys():
        w_overs = sum(o * r for o, r in t_overs[t_id]) / sum(r for _, r in t_overs[t_id]) if sum(r for _, r in t_overs[t_id]) > 0 else 0.0

        stats_list.append({
            "Team"              : t1_lookup.get(t_id, f"Team {t_id}"),
            "Median Vintage"    : format_year(np.median(t_vint  [t_id])),
            "Average GR"        : f"{np.mean(t_c_ps             [t_id])     * 100   :.2f}",
            "Rig Synergy"       : f"{np.mean(t_on_syn           [t_id])     * 100   :.2f}",
            "Off Synergy"       : f"{np.mean(t_off_syn          [t_id])     * 100   :.2f}",
            "Shared Rigs"       : f"{np.mean(t_sh_rig           [t_id])     * 100   :.2f}",
            "Total Solos"       : t_solos[t_id],
            "Average Overs"     : f"{w_overs                                        :.2f}"
        })

    df = pd.DataFrame(stats_list).sort_values("Average GR", ascending = False)
    export_df_to_png(df, png_dir, "Team.png", "Team Statistics")

def create_tier_report(s_part, raw_assignments, c_counts, p_pts, p_blks, png_dir):
    tiers           = sorted({v[1] for v in raw_assignments.values() if v[1] != "N/A"}, reverse = True)
    results         = []
    max_pts_seen    = -1
    max_blk_seen    = -1

    for tr in tiers:
        tp = [n for n in s_part if n.lower() in raw_assignments and raw_assignments[n.lower()][1] == tr]
        if not tp: continue
        
        def format_best(player_list, stat_dict):
            sorted_players = sorted(
                player_list,
                key     = lambda x: (stat_dict[x], c_counts[x] / s_part[x] if s_part[x] else 0),
                reverse = True
            )

            best_name   = sorted_players[0]
            val         = stat_dict[best_name]
            ties        = [p for p in player_list if stat_dict[p] == val]
            
            if len(ties) > 1:
                gr = f"{(c_counts[best_name] / s_part[best_name]) * 100:.2f}" if s_part[best_name] else "N/A"
                return f"{best_name} ({val}, {gr})"
            return f"{best_name} ({val})"

        best_atk_str = format_best(tp, p_pts)
        best_blk_str = format_best(tp, p_blks)
        
        cur_pts = p_pts     [sorted(tp, key = lambda x: p_pts[x],   reverse = True)[0]]
        cur_blk = p_blks    [sorted(tp, key = lambda x: p_blks[x],  reverse = True)[0]]

        atk_dis = best_atk_str if cur_pts > max_pts_seen else "▼"
        blk_dis = best_blk_str if cur_blk > max_blk_seen else "▼"

        results.append([tr, atk_dis, blk_dis])
        
        max_pts_seen = max(max_pts_seen, cur_pts)
        max_blk_seen = max(max_blk_seen, cur_blk)

    results.sort(key = lambda x: x[0])
    export_df_to_png(pd.DataFrame(results, columns = ["Tier", "Attacker", "Blocker"]), png_dir, "Tier.png", "Tier Bests")

def create_watched_report(s_part, p_l_corr, p_l_vint, png_dir):
    plist   = [n for n in s_part if p_l_corr[n]]
    e       = sorted(plist, key = lambda x: np.mean     (p_l_corr[x]), reverse = True)  [: 3]
    h       = sorted(plist, key = lambda x: np.mean     (p_l_corr[x]), reverse = False) [: 3]
    z       = sorted(plist, key = lambda x: np.median   (p_l_vint[x]), reverse = True)  [: 3]
    b       = sorted(plist, key = lambda x: np.median   (p_l_vint[x]), reverse = False) [: 3]
    
    rows = []
    for i in range(3):
        rows.append([f"{i+1}", 
            f"{e[i]} ({np.mean(p_l_corr[e[i]]):.2f})"               if i < len(e) else "N/A",
            f"{h[i]} ({np.mean(p_l_corr[h[i]]):.2f})"               if i < len(h) else "N/A",
            f"{z[i]} ({format_year(np.median(p_l_vint[z[i]]))})"    if i < len(z) else "N/A",
            f"{b[i]} ({format_year(np.median(p_l_vint[b[i]]))})"    if i < len(b) else "N/A"])
    export_df_to_png(pd.DataFrame(rows, columns = ["Rank", "Easiest", "Hardest", "Newest", "Oldest"]), png_dir, "Watched.png", "List Statistics")

if __name__ == "__main__":
    root = tk.Tk(); root.withdraw()
    process_files()