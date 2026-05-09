import  json
import  os
import  re
import  shutil
import  numpy       as      np
import  pandas      as      pd
import  tkinter     as      tk
from    collections import  Counter,    defaultdict
from    html2image  import  Html2Image
from    pathlib     import  Path
from    PIL         import  Image,      ImageChops,     ImageOps
from    tkinter     import  messagebox, ttk

BROWSER_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

DIR_DEPS        = "dependencies"
DIR_JSONS       = "jsons"
DIR_OUT         = "output"
DIR_TOURS       = "tours"
FILE_ALIASES    = "aliases.txt"
FILE_CODES      = "codes.txt"

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
    decimal     = next((val for s, val in season_map.items() if s in v_lower), 0.0)
    return year_val + decimal

def format_year(val):
    if val is None or val == "N/A": return "N/A"
    year    = int(val)
    frac    = val - year
    season  = "Winter" if frac < 0.25 else "Spring" if frac < 0.50 else "Summer" if frac < 0.75 else "Fall"
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

class UnifiedDialog(tk.Toplevel):
    def __init__(self, parent, title, prompt):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        main_frame = ttk.Frame(self, padding = 15)
        main_frame.pack(fill = tk.BOTH, expand = True)
        ttk.Label(main_frame, text = prompt, font = ("Segoe UI", 10)).pack(pady = (0, 10), anchor = "w")
        self.container = ttk.Frame(main_frame)
        self.container.pack(fill = tk.BOTH, expand = True)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill = tk.X, pady = (15, 0))
        self.confirm_btn = ttk.Button(btn_frame, text = "Confirm", command = self.on_confirm)
        self.confirm_btn.pack(side = tk.RIGHT, padx = 5)
        self.bind("<Return>", lambda: self.on_confirm())

    def on_confirm(self): self.destroy()

class StringDialog(UnifiedDialog):
    def __init__(self, parent, title, prompt, initialvalue = ""):
        super().__init__(parent, title, prompt)
        self.entry = ttk.Entry(self.container, width = 40)
        self.entry.insert(0, initialvalue)
        self.entry.pack(fill = tk.X)
        self.entry.focus_set()
        self.grab_set(); self.wait_window()

    def on_confirm(self):
        self.result = self.entry.get()
        super().on_confirm()

class SpinboxDialog(UnifiedDialog):
    def __init__(self, parent, title, prompt, initialvalue):
        super().__init__(parent, title, prompt)
        self.spin = ttk.Spinbox(self.container, from_ = 1, to = 6, width = 10)
        self.spin.set(initialvalue)
        self.spin.pack(anchor = "w")
        self.grab_set(); self.wait_window()

    def on_confirm(self):
        try                 : self.result = int(self.spin.get())
        except ValueError   : self.result = None
        super().on_confirm()

class NewPlayerDialog(UnifiedDialog):
    def __init__(self, parent, active_players):
        super().__init__(parent, "New Player Input", "Select new player(s), if any:")
        self.selected_new   = []
        self.blue_shade     = "#0056B3"
        self.vars           = {}
        player_list         = sorted    (list(active_players), key = str.lower)
        num_players         = len       (player_list)
        rows_per_col        = 8 if num_players >= 16 else num_players
        for i, name in enumerate(player_list):
            col             = i //  rows_per_col
            row             = i %   rows_per_col
            var             = tk.BooleanVar(value = False)
            self.vars[name] = var
            item_frame      = ttk.Frame(self.container)
            item_frame.grid(row = row, column = col, padx = 5, pady = 1, sticky = "w")
            box = tk.Canvas(item_frame, width = 10, height = 10, bg = "white", highlightthickness = 1, highlightbackground = "black")
            box.pack(side = tk.LEFT, padx = (0, 5))
            lbl = ttk.Label(item_frame, text = name, font = ("Segoe UI", 10))
            lbl.pack(side = tk.LEFT)
            for widget in (box, lbl): widget.bind("<Button-1>", lambda _, n=name, b=box: self.toggle_custom(n, b))
        self.grab_set(); self.wait_window()

    def toggle_custom(self, name, box):
        new_val = not self.vars[name].get()
        self.vars[name].set(new_val)
        color = self.blue_shade if new_val else "white"
        box.configure(bg = color)

    def on_confirm(self):
        self.selected_new = [name for name, var in self.vars.items() if var.get()]
        super().on_confirm()

class TourSelectionDialog(UnifiedDialog):
    def __init__(self, parent, tour_ids):
        super().__init__(parent, "Tour Selection", "Select tours to process:")
        self.selected_tours = []
        self.vars           = {}
        for _, tid in enumerate(tour_ids):
            var             = tk.BooleanVar(value = False)
            self.vars[tid]  = var
            ttk.Checkbutton(self.container, text = f"Tour {tid}", variable = var).pack(anchor = "w", pady = 2)
        self.grab_set(); self.wait_window()

    def on_confirm(self):
        self.selected_tours = [tid for tid, var in self.vars.items() if var.get()]
        super().on_confirm()

class ManualMatchDialog(tk.Toplevel):
    def __init__(self, parent, unknown_name, available_pool):
        super().__init__(parent)
        self.title("Manual Match Required")
        self.result = None
        ttk.Label(self, text = f"Match required for: '{unknown_name}'", font = ("Arial", 10, "bold")).pack(pady = 10)
        self.listbox = tk.Listbox(self, height = 15)
        self.listbox.pack(padx = 10, fill = tk.BOTH)
        for name in sorted(available_pool): self.listbox.insert(tk.END, name)
        ttk.Button(self, text = "Match Selected", command = self.on_match).pack(pady = 10)
        self.grab_set(); self.wait_window()

    def on_match(self):
        sel = self.listbox.curselection()
        if sel: self.result = self.listbox.get(sel[0]); self.destroy()

class SubSelectionDialog(tk.Toplevel):
    def __init__(self, parent, missing_roster):
        super().__init__(parent)
        self.title("Substitute Resolution")
        self.result = None
        tk.Label(self, text="Select the subbed player:").pack(padx = 20, pady = 10)
        self.listbox = tk.Listbox(self, height = len(missing_roster))
        self.listbox.pack(padx = 20, pady = 5, fill = tk.X)
        for m in missing_roster: self.listbox.insert(tk.END, m)
        ttk.Button(self, text = "Confirm", command = self.on_confirm).pack(pady = 10)
        self.grab_set(); self.wait_window()

    def on_confirm(self):
        sel = self.listbox.curselection()
        if sel: self.result = self.listbox.get(sel[0]); self.destroy()

# Main Processor
class TourAnalyzer:
    def __init__(self, tour_id):
        self.tour_id                    = str(tour_id)
        self.script_dir                 = Path(__file__).parent.absolute()
        self.tour_dir                   = self.script_dir / DIR_TOURS / self.tour_id
        self.browser_path               = self._find_browser()
        self.alias_map                  = self._load_aliases()
        self.s_part                     = defaultdict(int)
        self.c_counts                   = defaultdict(int)
        self.e_counts                   = defaultdict(int)
        self.p_rev_e                    = defaultdict(int)
        self.p_two_e                    = defaultdict(int)
        self.p_pts                      = defaultdict(int)
        self.p_blks                     = defaultdict(int)
        self.p_type_c                   = defaultdict(lambda: defaultdict(int))
        self.p_type_s                   = defaultdict(lambda: defaultdict(int))
        self.p_rigs                     = defaultdict(int)
        self.p_rigs_h                   = defaultdict(int)
        self.p_l_vint                   = defaultdict(list)
        self.p_l_corr                   = defaultdict(list)
        self.p_m_erigs                  = defaultdict(int)
        self.p_l_solos                  = defaultdict(int)
        self.p_chan_c                   = defaultdict(int)
        self.p_chan_s                   = defaultdict(int)
        self.t_vint                     = defaultdict(list)
        self.t_c_ps                     = defaultdict(list)
        self.t_on_syn                   = defaultdict(list)
        self.t_off_syn                  = defaultdict(list)
        self.t_sh_rig                   = defaultdict(list)
        self.t_solos                    = defaultdict(int)
        self.t_sweeps                   = defaultdict(int)
        self.t_overs                    = defaultdict(list)
        self.genre_c                    = Counter()
        self.tag_c                      = Counter()
        self.all_diff, self.all_vint    = [], []
        self.global_stats               = Counter()
        self.tour_label                 = ""
        self.chanting_ids               = set()

    def _find_browser(self): return next((p for p in BROWSER_PATHS if os.path.exists(p)), None)

    def _load_aliases(self):
        amap = {}
        path = self.script_dir / DIR_DEPS / FILE_ALIASES
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    if "," in line:
                        existing, new   = [x.strip() for x in line.split(",", 1)]
                        amap[new]       = existing
        return amap

    def _save_alias(self, existing, new):
        dep_dir = self.script_dir / DIR_DEPS
        dep_dir.mkdir(exist_ok = True)
        with open(dep_dir / FILE_ALIASES, "a", encoding = "utf-8") as f: f.write(f"{existing}, {new}\n")

    def run(self):
        chanting_path = self.script_dir / DIR_DEPS / "chanting" / "chanting.txt"
        if chanting_path.exists():
            with open(chanting_path, "r") as f:
                for line in f:
                    line = line.strip()
                    if line: self.chanting_ids.add(line)

        json_dir = self.tour_dir / DIR_JSONS
        if not json_dir.exists() or not any(json_dir.glob("*.json")):
            messagebox.showerror("Error", f"Folder not found or empty: {json_dir}")
            return

        json_paths                                                      = list(json_dir.glob("*.json"))
        all_known, appearances                                          = self._scan_players    (json_paths)
        use_teams, elo_map, assignments, t1_lookup, rosters, all_known  = self._load_team_data  (all_known)
        missing_list_count                                              = 0
        found_types                                                     = set()

        for path in json_paths:
            with open(path, encoding = "utf-8") as f: data = json.load(f)
            
            songs = data.get("songs", [])
            if not songs: continue

            raw_f_players = {p for s in songs for p in s.get("correctGuessPlayers", [])} | {ls["name"] for s in songs for ls in s.get("listStates", [])}
            final_members = set(raw_f_players)

            if use_teams:
                t_in_f = {assignments[p.lower()][0] for p in raw_f_players if p.lower() in assignments}
                for tid in t_in_f:
                    ros     = rosters[tid]
                    missing = [p for p in ros if p not in raw_f_players]
                    if len([p for p in ros if p in raw_f_players]) == 3 and missing:
                        res = SubSelectionDialog(None, missing).result if len(missing) > 1 else missing[0]
                        if res:
                            final_members.add(res)
                            potential_subs = list(raw_f_players - rosters[tid])
                            for sub_candidate in potential_subs:
                                if sub_candidate.lower() not in assignments: assignments[sub_candidate.lower()] = assignments[res.lower()]

                if len(final_members) < 8:
                    for tid in t_in_f: final_members.update(rosters[tid])

            apply_rev       = (len(final_members) % 2 == 0)
            max_s           = max(s.get("songNumber", 0) for s in songs)
            f_type_totals   = defaultdict(int)

            for song in songs:
                st = song.get("songInfo", {}).get("type")
                if st in [1, 2, 3]: 
                    f_type_totals   [st] += 1
                    found_types.add (st)

            for name in final_members:
                if name in raw_f_players:
                    self.s_part[name] += max_s
                    for t in [1, 2, 3]: self.p_type_s[name][t] += f_type_totals[t]

            for song in songs:
                si      = song      .get("songInfo", {})
                st      = si        .get("type")
                ann_id  = str(si    .get("annSongId"))
                is_chan = ann_id in self.chanting_ids
                
                if isinstance(si.get("animeGenre"), list): self.genre_c .update(si.get("animeGenre"))
                if isinstance(si.get("animeTags"),  list): self.tag_c   .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])
                
                correct                     =   set(song.get("correctGuessPlayers", []))
                ls                          =   song.get("listStates",              [])
                self.global_stats["tot_c"]  +=  len(correct)
                
                yr = extract_year(si.get("vintage"))
                if yr is not None                                       : self.all_vint.append(yr)
                if isinstance(si.get("animeDifficulty"), (int, float))  : self.all_diff.append(si.get("animeDifficulty"))
                if not ls                                               : missing_list_count += 1

                s_riggers = {p["name"] for p in ls}
                if len(ls) == 1:
                    u                   =   ls[0]["name"]
                    self.p_l_solos[u]   +=  1
                    if not (len(correct) == 1 and list(correct)[0] == u): self.p_m_erigs[u] += 1

                if use_teams:
                    t_list = list({assignments[p.lower()][0] for p in raw_f_players if p.lower() in assignments})
                    if len(t_list) == 2:
                        tA, tB = t_list[0],             t_list[1]
                        cA, cB = correct & rosters[tA], correct & rosters[tB]

                        if len(cA) == 4 and not cB: self.t_sweeps[tA] += 1; self.global_stats["sweeps"] += 1
                        if len(cB) == 4 and not cA: self.t_sweeps[tB] += 1; self.global_stats["sweeps"] += 1

                        for cur, opp in [(tA, tB), (tB, tA)]:
                            cC, oC = correct & rosters[cur], correct & rosters[opp]
                            if not oC: 
                                for p in cC: self.p_pts[p] += 1
                            if len(cC) == 1 and len(oC) > 0: self.p_blks[list(cC)[0]] += 1
                    
                    for tid in t_list:
                        ros     = rosters[tid]
                        c_on_t  = correct & ros
                        self.t_c_ps[tid].append(len(c_on_t) / 4.0)
                        if yr is not None: self.t_vint[tid].append(yr)
                        if s_riggers & ros:
                            self.t_on_syn       [tid].append(len(c_on_t)                / 4.0)
                            self.t_sh_rig       [tid].append((len(s_riggers & ros) - 1) / 3.0)
                            self.t_overs        [tid].append((len(correct), len(s_riggers & ros)))
                        else: self.t_off_syn    [tid].append(len(c_on_t)                / 4.0)

                if len(final_members - correct) == 0: self.global_stats["fulls"] += 1
                elif apply_rev and len(final_members - correct) == 1:
                    self.global_stats   ["sevens"]                          += 1
                    self.p_rev_e        [list(final_members - correct)[0]]  += 1
                elif len(correct) == 2:
                    self.global_stats   ["doubles"]     += 1
                    for p in correct: self.p_two_e[p]   += 1
                elif len(correct) == 1:
                    self.global_stats["solos"]                                              +=  1
                    sw                                                                      =   list(correct)[0]
                    self.e_counts[sw]                                                       +=  1
                    if sw.lower() in assignments: self.t_solos[assignments[sw.lower()][0]]  +=  1
                elif len(correct) == 0: self.global_stats["blanks"] += 1

                for name in final_members:
                    if name in correct:
                        self                        .c_counts[name]     += 1
                        if st in [1, 2, 3]  : self  .p_type_c[name][st] += 1
                        if is_chan          : self  .p_chan_c[name]     += 1
                    if is_chan: self                .p_chan_s[name]     += 1

                if ls:
                    for p in ls:
                        n = p["name"]       ; self.p_rigs   [n] += 1
                        if n in correct     : self.p_rigs_h [n] += 1
                        if yr is not None   : self.p_l_vint [n].append(yr)
                        self.p_l_corr[n].append(len(correct))

        self._finalize_outputs(missing_list_count, appearances, use_teams, elo_map, assignments, t1_lookup, found_types, all_known)

    def _scan_players(self, paths):
        players = set           ()
        apps    = defaultdict   (set)

        for p in paths:
            try:
                with open(p, encoding = "utf-8") as f:
                    data = json.load(f)
                    for s in data.get("songs", []):
                        for plyr    in s.get("correctGuessPlayers", []): players.add(plyr);         apps[plyr]          .add(str(p))
                        for ls      in s.get("listStates",          []): players.add(ls["name"]);   apps[ls["name"]]    .add(str(p))
            except: continue
        return players, apps

    def _load_team_data(self, all_known):
        codes = self.tour_dir / FILE_CODES
        if not codes.exists(): return False, {}, {}, {}, defaultdict(set)
        
        self.main_roster_names = set()
        elo_map, assignments, rosters, t1_lookup    = {}, {}, defaultdict(set), {}
        avail                                       = list(all_known)
        with open(codes, "r", encoding = "utf-8") as f: lines = f.readlines()

        for line in lines:
            matches = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', line)
            for p_in, val in matches:
                match = next((n for n in all_known if n.lower() == p_in.lower()), None)
                if not match and p_in in self.alias_map: match = next((n for n in all_known if n == self.alias_map[p_in]), None)
                if not match and ("[" in line or "Subs:" in line):
                    match = ManualMatchDialog(None, p_in, avail).result
                    if match: self._save_alias(match, p_in); self.alias_map[p_in] = match
                if match: elo_map[match.lower()] = val

        idx = 1
        for line in lines:
            if "[" not in line: continue
            pre     = re.match(r'^(?:\\s*)?([^:\[\d\(]+)\s*\([\d.-]+\):', line)
            ename   = pre.group(1).strip() if pre else None
            sec     = line.split(":", 1)[1] if ":" in line else line
            mems    = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', sec)

            for i, (p_in, _) in enumerate(mems[:4]):
                tier    = str(i + 1)
                match   = next((n for n in all_known if n.lower() == p_in.lower() or (p_in in self.alias_map and n == self.alias_map[p_in])), None)
                if match:
                    self.main_roster_names.add(match.lower())
                    assignments[match.lower()] = (idx, tier)
                    rosters[idx].add(match)
                    if match in avail: avail.remove(match)
                    t1_lookup[idx] = ename if ename else (match if tier == "1" else t1_lookup.get(idx))
            idx += 1
        return True, elo_map, assignments, t1_lookup, rosters, all_known

    def _finalize_outputs(self, missing_count, appearances, use_teams, elo_map, assignments, t1_lookup, found_types, original_roster):
        watched_valid       = missing_count <= 5
        baseline_initial    = int(np.median([len(appearances.get(name, [])) for name in self.s_part]))
        init_label          = "Watched" if watched_valid else "Usual"
        self.tour_label     = StringDialog(root, f"Name Input for Tour {self.tour_id}", "Enter the Tour name:", initialvalue = init_label).result
        if not self.tour_label: self.tour_label = init_label
        tour_disp           = f"{self.tour_label.strip()} Tour"
        global_dialog       = SpinboxDialog(root, f"Round Count for Tour {self.tour_id}", "Enter the expected amount of rounds:", baseline_initial)
        base_exp            = global_dialog.result
        if base_exp is None: base_exp = baseline_initial

        exp_map = {}
        for name in list(self.s_part.keys()):
            act = len(appearances.get(name, []))
            if act < base_exp:
                player_dialog   = SpinboxDialog(root, f"Count Mismatch Warning for Tour {self.tour_id}", f"Only {act} JSON(s) mention {name}; how many rounds were they expected to be in?", act)
                val             = player_dialog.result
                target          = val if val is not None else act
                exp_map[name]   = target

                if target > act:
                    avg_songs_per_json  =   sum(self.s_part.values()) / sum(len(v) for v in appearances.values())
                    missing_rounds      =   target - act
                    self.s_part[name]   +=  int(missing_rounds * avg_songs_per_json)
            else: exp_map[name] = base_exp

        new_players     = NewPlayerDialog(root, list(self.s_part.keys())).selected_new
        final_threshold = 6 if len(self.s_part) <= 20 else 5

        if      base_exp >= final_threshold     : stage = "Final"
        elif    base_exp == 3                   : stage = "Mid-Tour"
        else                                    : stage = f"R{base_exp}"

        type_map            = {1: "OP", 2: "ED", 3: "IN"}
        active_abbrs        = [type_map[t] for t in [1, 2, 3] if t in found_types]
        all_types_present   = set(type_map.keys()).issubset(found_types)
        
        if all_types_present    : type_str = ""
        else                    : type_str = f"{'-'.join(active_abbrs)} " if active_abbrs else ""

        prefix      = f"{tour_disp}, {type_str}"
        out_path    = self.script_dir / DIR_OUT / self.tour_id
        out_path.mkdir(parents = True, exist_ok = True)

        self._create_player_png (use_teams, elo_map, watched_valid, stage, out_path, appearances, prefix, exp_map, base_exp, assignments, new_players, t1_lookup, original_roster)
        self._create_tour_png   (use_teams, watched_valid, out_path)

        if watched_valid and assignments        : self._create_team_png     (t1_lookup,     out_path)
        if assignments                          : self._create_tier_png     (assignments,   out_path,   watched_valid)
        if watched_valid                        : self._create_watched_png  (out_path)
        if watched_valid and self.chanting_ids  : self._create_chanting_png (out_path)

        self._fuse_and_clean(out_path)
        messagebox.showinfo("Success", f"Saved Tour {self.tour_id} to {DIR_OUT}/{self.tour_id}")

    def _create_player_png(self, use_teams, elo_map, watched, stage, path, apps, prefix, exp_map, base_exp, assigns, new_players, t1_lookup, original_roster):
        rows, eligibility   = [], []
        t_labels            = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
        active              = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]

        for name in self.s_part:
            tot, cor    = self.s_part[name], self.c_counts[name]
            target      = exp_map.get(name, base_exp)
            d_name      = name

            if name in new_players: d_name += " ☆"

            if target < base_exp:
                if name.lower() in self.main_roster_names   : d_name += " ▼"
                else                                        : d_name += " ▲"
            
            is_eligible = not ("▼" in d_name or "▲" in d_name)
            eligibility.append(is_eligible)
            
            act = len(apps.get(name, []))
            if act < target:
                syms = ["", "①", "②", "③", "④", "⑤", "⑥"]
                if 0 < (target-act) < len(syms): d_name += f" {syms[target-act]}"

            row = {"Player": d_name}
            
            if use_teams:
                team_info = assigns.get(name.lower(), ("N/A", "N/A"))
                if team_info[0] != "N/A":
                    leader_name = t1_lookup.get(team_info[0], "")
                    clean_name  = "".join(filter(str.isalnum, leader_name))
                    row["Team"] = clean_name[:3].upper() if leader_name else f"T{team_info[0]}"
                    row["Tier"] = team_info[1]
                else:
                    row["Team"] = "N/A"
                    row["Tier"] = "N/A"
                row["Elo"]      = elo_map.get(name.lower(), "N/A")

            row.update({
                "Guess Rate"    : cor / tot if tot else 0,
                "1/8s"          : self.e_counts [name],
                "2/8s"          : self.p_two_e  [name],
                "7/8s"          : self.p_rev_e  [name]
            })

            if use_teams: row.update({"Lives Taken": self.p_pts[name], "Lives Saved": self.p_blks[name]})
            
            for tid in active:
                seen                = self.p_type_s[name][tid]
                row[t_labels[tid]]  = self.p_type_c[name][tid] / seen if seen else np.nan

            if watched:
                row.update({
                    "Rigs"              : self.p_rigs[name],
                    "Rig Rate"          : self.p_rigs[name]             / tot                       if tot                          else np.nan,
                    "Rig Delta"         : (cor - self.p_rigs[name])     / cor                       if cor                          else np.nan,
                    "Rig GR"            : self.p_rigs_h[name]           / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Off GR"            : (cor - self.p_rigs_h[name])   / (tot - self.p_rigs[name]) if (tot - self.p_rigs[name])    else np.nan,
                })

            rows.append(row)

        df      = pd.DataFrame(rows).sort_values("Guess Rate", ascending = False)
        mask    = pd.Series(eligibility, index = pd.DataFrame(rows).index).reindex(df.index).values
        pcts    = ["Guess Rate"] + [t_labels[t] for t in active] + (["Rig Rate", "Rig Delta", "Rig GR", "Off GR"] if watched else [])

        if "Elo" in df.columns: df["Elo"] = pd.to_numeric(df["Elo"], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        for c in pcts: df[c] = pd.to_numeric(df[c], errors = 'coerce').mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        self._export_png(df, path, "Player.png", f"{prefix}Player Statistics, {stage}", mask)

    def _create_tour_png(self, use_teams, watched, path):
        def fmt_most(names, val):
            if not names: return "N/A"
            win = sorted(names, key = lambda x: (self.c_counts[x] / self.s_part[x]) if self.s_part[x] else 0)[0]
            gr  = (self.c_counts[win] / self.s_part[win]) * 100 if self.s_part[win] else 0
            return f"{win} ({val}{f', {gr:.2f}' if len(names) > 1 else ''})"

        stats = [
            ["Median Vintage",      format_year(round(np.median(self.all_vint), 2))                         if self.all_vint    else "N/A"],
            ["Average Difficulty",  f"{np.mean(self.all_diff):.2f}"                                         if self.all_diff    else "N/A"],
            ["Average GR",          f"{100 * (self.global_stats['tot_c'] / sum(self.s_part.values())):.2f}" if self.s_part      else "0.00"],
            ["Total 0/8s",          self.global_stats["blanks"]],
            ["Total 1/8s",          self.global_stats["solos"]],
            ["Total 2/8s",          self.global_stats["doubles"]],
            ["Total 7/8s",          self.global_stats["sevens"]],
            ["Total 8/8s",          self.global_stats["fulls"]]
        ]

        if use_teams: stats.append(["Total 4-0s", self.global_stats["sweeps"]])
        stats.extend([
            ["Most Popular Genre",  f"{self.genre_c .most_common(1)[0][0]} ({self.genre_c   .most_common(1)[0][1]})" if self.genre_c else "N/A"],
            ["Most Popular Tag",    f"{self.tag_c   .most_common(1)[0][0]} ({self.tag_c     .most_common(1)[0][1]})" if self.tag_c else "N/A"],
            ["Most 1/8s",           fmt_most([n for n, v in self.e_counts   .items() if v == max(self.e_counts  .values(), default = 0) and v > 0], max(self.e_counts   .values(), default = 0))],
            ["Most 2/8s",           fmt_most([n for n, v in self.p_two_e    .items() if v == max(self.p_two_e   .values(), default = 0) and v > 0], max(self.p_two_e    .values(), default = 0))],
            ["Most 7/8s",           fmt_most([n for n, v in self.p_rev_e    .items() if v == max(self.p_rev_e   .values(), default = 0) and v > 0], max(self.p_rev_e    .values(), default = 0))]
        ])
        
        plist   = list(self.s_part.keys())
        no_s    = sorted([n for n in plist if self.e_counts[n] ==   0 and self.s_part[n] > 0], key = lambda x: self.c_counts[x] / self.s_part[x], reverse = True)
        yes_s   = sorted([n for n in plist if self.e_counts[n] >    0 and self.s_part[n] > 0], key = lambda x: self.c_counts[x] / self.s_part[x])

        if no_s     : stats.append(["Highest GR Without 1/8s",  f"{no_s     [0]} ({100 * (self.c_counts[no_s    [0]] / self.s_part[no_s     [0]]):.2f})"])
        if yes_s    : stats.append(["Lowest GR With 1/8s",      f"{yes_s    [0]} ({100 * (self.c_counts[yes_s   [0]] / self.s_part[yes_s    [0]]):.2f}, {self.e_counts[yes_s[0]]})"])

        if watched:
            conv        = []
            eligible    = [p for p in plist if self.p_l_solos[p] > 0]
            
            if eligible:
                total_hits      = sum((self.p_l_solos[p] - self.p_m_erigs[p]) for p in eligible)
                total_attempts  = sum(self.p_l_solos[p] for p in eligible)
                global_avg      = total_hits / total_attempts if total_attempts > 0 else 0
                constant        = 3
                
                for n in eligible:
                    t               = self.p_l_solos[n]
                    h               = t - self.p_m_erigs[n]
                    weighted_score  = (h + constant * global_avg) / (t + constant)
                    conv.append({'n': n, 'score': weighted_score, 'p': 100 * h / t, 'h': h, 't': t})

                b = sorted(conv, key = lambda x: x['score'], reverse = True)    [0]
                w = sorted(conv, key = lambda x: x['score'])                    [0]

                stats.append(["Best Solo Rig Converter",    f"{b['n']} ({b['p']:.2f}%, {b['h']}/{b['t']})"])
                stats.append(["Worst Solo Rig Converter",   f"{w['n']} ({w['p']:.2f}%, {w['h']}/{w['t']})"])

        self._export_png(pd.DataFrame(stats, columns = ["Statistic", "Value"]), path, "Tour.png", "Tour Statistics")

    def _create_team_png(self, t1_lookup, path):
        res = []
        for tid in self.t_c_ps:
            res.append({
                "Team Leader"       : t1_lookup.get(tid, f"Team {tid}"),
                "Median Vintage"    : format_year(np.median(self.t_vint[tid])),
                "Average GR"        : f"{np.mean(self.t_c_ps    [tid]) * 100:.2f}",
                "Rig Synergy"       : f"{np.mean(self.t_on_syn  [tid]) * 100:.2f}",
                "Off Synergy"       : f"{np.mean(self.t_off_syn [tid]) * 100:.2f}",
                "Shared Rigs"       : f"{np.mean(self.t_sh_rig  [tid]) * 100:.2f}",
                "Total 1/8s"        : self.t_solos[tid],
            })
        self._export_png(pd.DataFrame(res).sort_values("Average GR", ascending = False), path, "Team.png", "Team Statistics")

    def _create_tier_png(self, assigns, path, watched_valid):
        tiers   = sorted({v[1] for v in assigns.values() if v[1] != "N/A"}, reverse = True)
        res     = []

        for tr in tiers:
            tp = [n for n in self.s_part if n.lower() in assigns and assigns[n.lower()][1] == tr]
            if not tp: continue

            def get_generalist(plist):
                sorted_p        = sorted(plist, key = lambda x: (self.c_counts[x] / self.s_part[x] if self.s_part[x] else 0), reverse = True)
                name, value     = sorted_p[0], 100 * (self.c_counts[sorted_p[0]] / self.s_part[sorted_p[0]]) if self.s_part[sorted_p[0]] else 0
                return f"{name} ({value:.2f})"

            def get_attblk(plist, sdict):
                sorted_p        = sorted(plist, key = lambda x: (sdict[x], self.c_counts[x] / self.s_part[x] if self.s_part[x] else 0), reverse = True)
                name, value     = sorted_p[0], sdict[sorted_p[0]]
                return f"{name} ({value})"

            def get_contributor(plist):
                sorted_p        = sorted(plist, key = lambda x: ((self.p_pts[x] + self.p_blks[x]), (self.c_counts[x] / self.s_part[x] if self.s_part[x] else 0)), reverse = True)
                name, v1, v2    = sorted_p[0], self.p_pts[sorted_p[0]], self.p_blks[sorted_p[0]]
                return f"{name} ({v1}, {v2})"

            def get_chanter(plist):
                pool = [n for n in plist if self.p_chan_s[n] > 0 and self.c_counts[n] > 0]
                if not pool: return "N/A"
                sorted_p    = sorted(pool, key = lambda x: (100 * self.p_chan_c[x] / self.p_chan_s[x], -(100 * self.c_counts[x] / self.s_part[x])), reverse = True)
                name        = sorted_p[0]                
                ratio       = 100 * self.p_chan_c[name] / self.p_chan_s[name]
                return f"{name} ({ratio:.2f})"

            row_data = [tr, get_generalist(tp), get_attblk(tp, self.p_pts), get_attblk(tp, self.p_blks), get_contributor (tp)]
            if watched_valid: row_data.append(get_chanter(tp))
            res.append(row_data)
            
        cols = ["Tier", "Generalist", "Attacker", "Blocker", "Contributor"]
        if watched_valid: cols.append("Chanter")
        self._export_png(pd.DataFrame(sorted(res, key = lambda x: x[0]), columns = cols), path, "Tier.png", "Tier Bests")

    def _create_watched_png(self, path):
        plist   = [n for n in self.s_part if self.p_l_corr[n]]
        e       = sorted(plist, key = lambda x: np.mean     (self.p_l_corr[x]), reverse = True) [ : 3]
        h       = sorted(plist, key = lambda x: np.mean     (self.p_l_corr[x]))                 [ : 3]
        z       = sorted(plist, key = lambda x: np.median   (self.p_l_vint[x]), reverse=True)   [ : 3]
        b       = sorted(plist, key = lambda x: np.median   (self.p_l_vint[x]))                 [ : 3]

        rows    = [[
            f"{i+1}", 
            f"{e[i]} ({np.mean(self.p_l_corr[e[i]]):.2f})"              if i < len(e) else "N/A",
            f"{h[i]} ({np.mean(self.p_l_corr[h[i]]):.2f})"              if i < len(h) else "N/A",
            f"{z[i]} ({format_year(np.median(self.p_l_vint[z[i]]))})"   if i < len(z) else "N/A",
            f"{b[i]} ({format_year(np.median(self.p_l_vint[b[i]]))})"   if i < len(b) else "N/A"
        ] for i in range(3)]

        self._export_png(pd.DataFrame(rows, columns = ["Rank", "Easiest", "Hardest", "Newest", "Oldest"]), path, "Watched.png", "List Statistics")

    def _create_chanting_png(self, path):
        plist = [n for n in self.s_part if self.p_chan_s[n] > 0 and self.c_counts[n] > 0]
        if not plist: return

        def get_ratio   (p): return 100 * self.p_chan_c[p] / self.p_chan_s  [p]
        def get_gr      (p): return 100 * self.c_counts[p] / self.s_part    [p]

        best    = sorted(plist, key = lambda p: (get_ratio(p), -get_gr(p)), reverse = True) [ : 3]
        worst   = sorted(plist, key = lambda p: (get_ratio(p), -get_gr(p)))                 [ : 3]
        rows    = []

        for i in range(3):
            b_cell = "N/A"
            if i < len(best):
                p       = best[i]
                b_cell  = f"{p} ({get_ratio(p):.2f})"
            
            w_cell = "N/A"
            if i < len(worst):
                p       = worst[i]
                w_cell  = f"{p} ({get_ratio(p):.2f})"
            
            rows.append([f"{i + 1}", b_cell, w_cell])

        self._export_png(pd.DataFrame(rows, columns = ["Rank", "Best", "Worst"]), path, "Chanting.png", "Chanting Statistics")

    def _export_png(self, df, path, fname, title, mask = None):
        if not self.browser_path: return

        desc    = ["Elo", "Guess Rate", "1/8s", "2/8s", "Rigs", "Rig Delta", "Lives Taken", "Lives Saved", "Rig Rate", "OP GR", "ED GR", "IN GR", "Rig GR", "Off GR", "Average GR", "Rig Synergy", "Off Synergy", "Shared Rigs", "Total 1/8s"]
        asc     = ["7/8s"]
        rest    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs"]
        stats   = {}

        for col in df.columns:
            if col in desc or col in asc:
                num     = pd.to_numeric(df[col].astype(str).str.replace('%',''), errors = 'coerce')
                el_num  = num[mask].dropna() if mask is not None and col in rest else num.dropna()
                if not num.dropna().empty:
                    stats[col] = {
                        'max'   : num.dropna().max(),
                        'min'   : el_num.min() if not el_num.empty else None, 
                        's_max' : num.dropna().value_counts().get(num.dropna().max(), 0) <= 3, 
                        's_min' : el_num.value_counts().get(el_num.min(), 0) <= 3 if not el_num.empty else False
                    }

        df      = df.reset_index(drop = True)
        borders = []

        if "Guess Rate" in df.columns:
            if      self.tour_label == "Watched 2+8"            : init_val = "25, 20, 15, 10, 5"
            elif    self.tour_label == "Watched"                : init_val = "28, 18, 12, 6"
            elif    self.tour_label in ["Usual", "Quagsual"]    : init_val = "28, 19, 8"
            elif    "Rigs"          in df.columns               : init_val = "28, 18, 12, 6"
            else                                                : init_val = "28, 19, 8"
                
            trimmed_title   = title.split(" Tour")[0]
            val_str         = StringDialog(root, f"Threshold Input for Tour {self.tour_id}", f"Enter comma-separated thresholds for {trimmed_title}:", initialvalue = init_val).result

            try         : th = [float(x.strip()) for x in val_str.split(",")] if val_str else []
            except      : th = [28.0, 18.0, 12.0, 6.0]
            
            gv = pd.to_numeric(df["Guess Rate"].astype(str).str.replace('%',''), errors = 'coerce').tolist()

            for t in th:
                f_idx = -1
                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i
                if f_idx != -1 and f_idx < len(df) - 1: borders.append(f_idx)

        html = f"<thead><tr>" + "".join([f"<th>{str(c).replace(' ','<br>')}</th>" for c in df.columns]) + "</tr></thead><tbody>"
        for idx, row in df.iterrows():
            b_s     =   "border-bottom: 3px solid black;" if idx in borders else ""
            html    +=  "<tr>"

            for i, (cname, cell) in enumerate(row.items()):
                style = [b_s] if b_s else []
                if cname in stats:
                    v = pd.to_numeric(str(cell).replace('%',''), errors = 'coerce')
                    if pd.notnull(v):
                        is_max, is_min  = (v == stats[cname]['max']) and stats[cname]['s_max'], (v == stats[cname]['min']) and stats[cname]['s_min']
                        elig            = True if mask is None or cname not in rest else mask[idx]
                        if cname in desc:
                            if      is_max          : style.append("color: #0056B3; font-weight: bold;")
                            elif    is_min and elig : style.append("color: #D95400; font-weight: bold;")
                        elif cname in asc:
                            if      is_max and elig : style.append("color: #D95400; font-weight: bold;")
                            elif    is_min          : style.append("color: #0056B3; font-weight: bold;")
                s_attr  =   f' style="{" ".join(style)}"' if style else ""
                cnt     =   f"<b>{cell}</b>" if i==0 else cell
                html    +=  f"<td{s_attr}>{cnt}</td>"
            html += "</tr>"

        full    = f"<html><head><style>body {{font-family: 'Segoe UI', Arial, sans-serif; background: white; display: inline-block; margin: 0;}} h2 {{margin: 10px 0 10px 5px; font-size: 30px; text-align: center;}} table {{margin-left: 10px; border-collapse: collapse; width: auto;}} th {{font-weight: bold; font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}} td {{font-size: 20px; text-align: center; padding: 10px; border: 1px solid black;}}</style></head><body><h2>{title}</h2><table>{html}</table></body></html>"
        hti     = Html2Image(size = (max(2000, len(df.columns) * 120), max(2000, len(df) * 60)), browser_executable = self.browser_path, output_path = str(path), custom_flags = ['--log-level=3', '--silent'])

        hti.screenshot(html_str = full, save_as = fname)
        try     : trim_whitespace(path / fname)
        except  : pass

    def _fuse_and_clean(self, path):
        f       = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "Watched": "Watched.png", "Chanting": "Chanting.png"}
        ps      = {k: path / v      for k, v in f   .items() if (path / v).exists()}
        imgs    = {k: Image.open(v) for k, v in ps  .items()}
        if not imgs: return

        if "Team" in imgs and "Chanting" in imgs:
            t_img, c_img    = imgs["Team"], imgs["Chanting"]
            combined_w      = t_img.width + 10 + c_img.width
            combined_h      = max(t_img.height, c_img.height)
            combined        = Image.new("RGB", (combined_w, combined_h), "white")

            combined.paste(t_img, (0, 0))
            combined.paste(c_img, (t_img.width + 10, 0))
            
            imgs["Team"] = combined
            del imgs["Chanting"]

        rk = [k for k in ["Team", "Tier", "Watched", "Chanting"] if k in imgs]
        if not rk: fused = imgs.get("Tour")
        elif len(rk) == 1:
            tw, th  = (imgs["Tour"].width, imgs["Tour"].height) if "Tour" in imgs else (0, 0)
            ok      = rk[0]
            fused   = Image.new("RGB", (max(tw, imgs[ok].width), th + (10 if th else 0) + imgs[ok].height), "white")
            if "Tour" in imgs: fused.paste(imgs["Tour"], (0, 0))
            fused.paste(imgs[ok], (0, th + 10 if th else 0))
        else:
            tw, th  = (imgs["Tour"].width, imgs["Tour"].height) if "Tour" in imgs else (0, 0)
            rw, rh  = max([imgs[k].width for k in rk]), sum([imgs[k].height + 10 for k in rk]) - 10
            fused   = Image.new("RGB", (tw + (10 if tw and rw else 0) + rw, max(th, rh)), "white")
            if "Tour" in imgs: fused.paste(imgs["Tour"], (0, 0))
            cx, cy  = (tw + 10 if tw else 0), 0
            for k in rk: fused.paste(imgs[k], (cx, cy)); cy += imgs[k].height + 10
            
        if fused:
            f_p = path / "Extra.png"
            fused.save(f_p)
            try     : trim_whitespace(f_p)
            except  : pass
            
        for p in ps.values():
            try     : os.remove(p)
            except  : pass

if __name__ == "__main__":
    root                = tk.Tk(); root.withdraw()
    selection_dialog    = TourSelectionDialog(root, ["0", "1", "2"])
    selected_tours      = selection_dialog.selected_tours
    
    if selected_tours:
        for tour_id in selected_tours: TourAnalyzer(tour_id).run()