import datetime, gspread, hashlib, json, logging, math, matplotlib, os, re, shutil, subprocess, sys, time

logging     .getLogger  ("adjustText").setLevel(logging.ERROR)
matplotlib  .use        ('Agg')

import concurrent.futures       as fut
import matplotlib.colors        as mc
import matplotlib.pyplot        as plt
import numpy                    as np
import pandas                   as pd

from adjustText             import adjust_text
from bs4                    import BeautifulSoup
from collections            import Counter, defaultdict
from curl_cffi              import requests
from dateutil.relativedelta import relativedelta
from hako.help.config       import *
from hako.help.dialog       import *
from html2image             import Html2Image
from pathlib                import Path
from PIL                    import Image
from scipy.spatial          import ConvexHull
from tkinter                import messagebox
from urllib.parse           import urlparse, urlunparse

TEAMS_RE = r"([^\s(]+)\s*\(([-]?\d+(?:\.\d+)?)\)"

def _nested_int_defaultdict     (): return defaultdict(int)
def _nested_list_defaultdict    (): return defaultdict(list)

class TourAnalyzer:
    def __init__(self, tour_id):
        self.tour_id                    = str(tour_id)
        self.script_dir                 = Path(__file__).parent.absolute()
        self.tour_dir                   = self.script_dir / DIR_TOURS / self.tour_id
        self.browser_path               = self._find_browser()
        self.s_part                     = defaultdict(int)
        self.c_counts                   = defaultdict(int)
        self.e_counts                   = defaultdict(int)
        self.p_rev_e                    = defaultdict(int)
        self.p_two_e                    = defaultdict(int)
        self.p_pts                      = defaultdict(int)
        self.p_blks                     = defaultdict(int)
        self.p_type_c                   = defaultdict(_nested_int_defaultdict)
        self.p_type_s                   = defaultdict(_nested_int_defaultdict)
        self.p_rigs                     = defaultdict(int)
        self.p_rigs_h                   = defaultdict(int)
        self.p_l_vint                   = defaultdict(list)
        self.p_c_vint                   = defaultdict(list)
        self.p_l_corr                   = defaultdict(list)
        self.p_lh_vint                  = defaultdict(list)
        self.p_lh_corr                  = defaultdict(list)
        self.p_m_erigs                  = defaultdict(int)
        self.p_l_solos                  = defaultdict(int)
        self.p_hit_diff                 = defaultdict(list)
        self.p_hit_vint                 = defaultdict(list)
        self.p_chan_c                   = defaultdict(int)
        self.p_chan_s                   = defaultdict(int)
        self.p_usefulness_sum           = defaultdict(float)
        self.p_overs_sum                = defaultdict(int)
        self.p_answer_times             = defaultdict(list)
        self.t_vint                     = defaultdict(list)
        self.t_c_ps                     = defaultdict(list)
        self.t_on_syn                   = defaultdict(list)
        self.t_off_syn                  = defaultdict(list)
        self.t_sh_rig                   = defaultdict(list)
        self.t_solos                    = defaultdict(int)
        self.t_sweeps                   = defaultdict(int)
        self.genre_c                    = Counter()
        self.tag_c                      = Counter()
        self.global_stats               = Counter()
        self.all_diff, self.all_vint    = [], []
        self.song_history               = []
        self.song_data                  = []
        self.chanting_ids               = set()
        self.subbed_players_set         = set()
        self.tour_label                 = ""
        self.id_database                = {}
        self.player_acronyms            = {}
        self.player_song_details        = defaultdict(_nested_list_defaultdict)
        self.tour_song_details          = defaultdict(list)
        self.team_song_details          = defaultdict(_nested_list_defaultdict)
        self.matrix_song_details        = defaultdict(list)
        self.raw_vintage_by_guess       = defaultdict(list)
        self.raw_vintage_by_list        = defaultdict(list)

    def _find_browser(self): return next((p for p in BROWSER_PATHS if os.path.exists(p)), None)

    def _load_player_ids(self):
        id_map = {}

        try:
            df = pd.read_csv(URL_ALIAS)

            for _, row in df.iterrows():
                name    = str(row.get('Player Name',    '')).strip().lower()
                pid     = str(row.get('Player ID',      '')).strip()

                if name and pid: id_map[name] = pid

        except: pass

        return id_map

    def _internal_clean_data(self, idtable, statstable, isWatched):
        headers                 = idtable[0]
        data                    = idtable[1:]
        alias_df                = pd.DataFrame(data, columns = headers)
        alias_df["Player Name"] = alias_df["Player Name"].str.strip().str.lower()
        alias_to_id             = dict(zip(alias_df["Player Name"], alias_df["Player ID"]))

        headers = statstable[0]
        data    = statstable[1:]
        df      = pd.DataFrame(data, columns = headers)
        df      = df.replace(r"^\s*$", pd.NA, regex = True).dropna(how = "all")
        
        df["Player ID"] = df["Player name"].dropna().str.strip().str.lower().map(alias_to_id)
        df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors = "coerce")
        
        cols = [
            "Rank",
            "Guess rate",
            "Usefulness",
            "erigs",
            "7/8s",
            "avg/8",
            "Lives taken",
            "Lives saved", 
            "WIN",
            "LOSE",
            "TIE",
            "Total hit",
            "OP guess rate",
            "ED guess rate",
            "IN guess rate"
        ]

        watched_cols = [
            "Rigs hit",
            "Rigs",
            "Rigs missed",
            "Solo rigs", 
            "Missed solos",
            "Lives lost on rigs",
            "Offlist erigs",
            "avg/8 of your rigs"
        ]
        
        df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")
        
        if isWatched:
            df[watched_cols]    = df[watched_cols].apply(pd.to_numeric, errors="coerce")
            cols.extend(watched_cols)
            df["Offlist hit"]   = df["Total hit"] - df["Rigs hit"]
        
        df = df[(
            pd.to_numeric(df["WIN"],    errors = 'coerce').fillna(0) + 
            pd.to_numeric(df["LOSE"],   errors = 'coerce').fillna(0) + 
            pd.to_numeric(df["TIE"],    errors = 'coerce').fillna(0)
        ) >= 4]

        return df

    def _clean_data_local(self, idtable, statstable, maxFallbackWindow, activeTours, is_list):
        df = self._internal_clean_data(idtable, statstable, is_list)
        if df.empty: return pd.DataFrame(columns = ["Player ID", "Guess rate", "Usefulness", "OP guess rate", "ED guess rate", "IN guess rate"])

        six_months_ago  = datetime.datetime.now() - relativedelta(months=maxFallbackWindow)
        year_6m_ago     = six_months_ago.year
        month_6m_ago    = six_months_ago.month

        year_df = df[((df["Timestamp"].dt.year > year_6m_ago)) | ((df["Timestamp"].dt.year == year_6m_ago) & (df["Timestamp"].dt.month >= month_6m_ago))]
        if year_df.empty: return pd.DataFrame(columns=["Player ID", "Guess rate", "Usefulness", "OP guess rate", "ED guess rate", "IN guess rate"])

        year_df     = year_df.sort_values(["Player ID", "Timestamp"])
        result_df   = year_df.groupby("Player ID").tail(activeTours)
        
        GR          = ["Guess rate", "Usefulness", "OP guess rate", "ED guess rate", "IN guess rate"]
        agg_dict    = {col: "mean" if col in GR else "max" for col in result_df.columns if col != "Player ID"}
        agg_dict    = {k: v for k, v in agg_dict.items() if k in result_df.columns}

        result_df = result_df.groupby("Player ID").agg(agg_dict).reset_index()
        result_df["Player ID"] = result_df["Player ID"].astype(int)

        return result_df

    def _generate_acronyms(self, active_names):
        acronyms = {}

        for name in active_names:
            clean                   = "".join(filter(str.isalnum, name))
            length                  = 3
            acr                     = clean[:length].upper() if len(clean) >= length else clean.upper().ljust(length, 'X')
            acronyms[name.lower()]  = acr

        while True:
            counts      = Counter(acronyms.values())
            duplicates  = {acr for acr, count in counts.items() if count > 1}

            if not duplicates: break

            for name in active_names:
                n_low = name.lower()

                if acronyms[n_low] in duplicates:
                    clean       = "".join(filter(str.isalnum, name))
                    curr_len    = len(acronyms[n_low])
                    next_len    = curr_len + 1

                    if next_len <= len(clean)   : acronyms[n_low] = clean[:next_len].upper()
                    else                        : acronyms[n_low] = clean.upper() + str(next_len - len(clean))

        self.player_acronyms = acronyms

    def _get_player_acronym(self, name): return self.player_acronyms.get(name.lower(), name[ : 3].upper())

    def _get_team_acronym(self, leader_name, tid):
        if leader_name:
            clean = "".join(filter(str.isalnum, leader_name))
            return clean[ : 3].upper()

        return f"T{tid}"

    def prepare_configuration(self):
        chanting_path = self.script_dir / DIR_TOURS / FILE_CHANT

        if chanting_path.exists():
            with open(chanting_path, "r") as f:
                for line in f:
                    line = line.strip()
                    if line: self.chanting_ids.add(line)

        json_dir = self.tour_dir / DIR_JSONS

        if not json_dir.exists() or not any(json_dir.glob("*.json")):
            messagebox.showerror("Error", f"Folder not found or empty: {json_dir}")
            return False

        self.json_paths = list(json_dir.glob("*.json"))
        fingerprints    = defaultdict(list)

        for path in self.json_paths:
            try:
                with open(path, encoding = "utf-8") as f: data = json.load(f)

            except Exception as e:
                messagebox.showerror("JSON corrupted", f"Could not read {path.name}: {e}")
                return False

            songs = data.get("songs", [])

            if not path.stem.startswith("amq"):
                match_digits = re.search(r'(\d+)$', path.stem)

                if match_digits:
                    m = int(match_digits.group(1))

                    if m <= THRESH_SONG:
                        songs           = songs[:min(m, len(songs))]
                        data["songs"]   = songs

            if not isinstance(songs, list) or not songs:
                messagebox.showerror("Disconnected Player JSON", f"Error in {path.name}: The exporter likely disconnected; ask someone else to re-upload this JSON")
                return False

            for _, song in enumerate(songs, 1):
                if not isinstance(song, dict) or "videoUrl" not in song:
                    messagebox.showerror("Disconnected Player JSON", f"Error in {path.name}: The exporter likely disconnected; ask someone else to re-upload this JSON")
                    return False

            payload = json.dumps(songs, sort_keys = True, ensure_ascii = False, default = str)
            fp      = hashlib.sha256(payload.encode("utf-8")).hexdigest()

            fingerprints[fp].append(path.name)

        duplicates = [names for names in fingerprints.values() if len(names) > 1]

        if duplicates:
            msg = "\n".join([f"• {', '.join(set_files)}" for set_files in duplicates])
            messagebox.showerror("Identical JSONs", f"The following files are from the same match:\n{msg}\nDelete one")
            return False

        all_known, self.apps = self._scan_players(self.json_paths)
        self._generate_acronyms(all_known)
        
        loaded = self._load_team_data(all_known)
        self.use_teams, self.elo_map, self.assignments, self.t1_lookup, self.rosters, all_known, sub_candidates_raw, original_players_display = loaded
        if not loaded[0]: self.use_teams = False

        self.missing_list_count = 0
        self.tour_types         = set()

        for path in self.json_paths:
            with open(path, encoding = "utf-8") as f: data = json.load(f)
            songs = data.get("songs", [])
            if not songs: continue

            for song in songs:
                st = song.get("songInfo", {}).get("type")

                if st in [1, 2, 3]                  : self.tour_types.add(st)
                if not song.get("listStates", [])   : self.missing_list_count += 1

        total_jsons         = len(self.json_paths)
        total_players       = len(all_known)
        baseline_initial    = total_jsons // (2 if total_players <= THRESH_PLYR else 3)
        watched_valid       = self.missing_list_count <= THRESH_WTCH

        if len(self.tour_types) == 1:
            t_map       = {1: "OP", 2: "ED", 3: "IN"}
            t_str       = t_map.get(list(self.tour_types)[0], "")
            init_label  = f"Watched {t_str}" if watched_valid else f"Random {t_str}"

        else: init_label = "Watched" if watched_valid else "Usual"

        if "Eru" in init_label and self.use_teams: default_th = ""

        else:
            if      init_label == "Watched 2+8s"                : default_th = "25, 20, 15, 10, 5"
            elif    init_label in ["Watched",   "QuagWatched"]  : default_th = "28, 18, 12, 6"
            elif    init_label in ["Usual",     "Quagsual"]     : default_th = "28, 19, 8"
            else                                                : default_th = "28, 19, 8"

        meta_dialog = TourMetadataDialog(None, self.tour_id, init_label, default_th, baseline_initial, list(all_known), self.elo_map, sub_candidates_raw, original_players_display, self.tour_dir)
        if meta_dialog.result is None: sys.exit(0)

        meta_res                = meta_dialog.result
        self.tour_label         = meta_res["tour_label"]
        self.delta_choice       = meta_res.get("delta_choice",      "No")
        self.challonge_choice   = meta_res.get("challonge_choice",  "No")
        
        if self.delta_choice == "Yes" and self.tour_label in TOUR_MODE_SHEET_MAP:
            print(f"[?] Fetching historic baselines")

            cred_file = self.script_dir / "help" / DIR_CREDS / "credentials.json"
            auth_file = self.script_dir / "help" / DIR_CREDS / "authorized_user.json"

            try:
                gc          = gspread.oauth(credentials_filename = str(cred_file), authorized_user_filename = str(auth_file))
                sheet       = gc.open_by_key(NGM_STATS_SHEET_ID)
                wks_ids     = sheet.get_worksheet_by_id(SHEET_PLAYER_IDS)
                rows_ids    = wks_ids.get_all_values()
                sheet_ref   = TOUR_MODE_SHEET_MAP[self.tour_label]

                if isinstance(sheet_ref, int)   : wks_stats = sheet.get_worksheet_by_id(sheet_ref)
                else                            : wks_stats = sheet.worksheet(sheet_ref)

                rows_stats          = wks_stats.get_all_values()
                is_list_mode        = "Watched" in self.tour_label or self.tour_label == "Watched"
                avg_df              = self._clean_data_local(rows_ids, rows_stats, 6, 10, is_list_mode)
                history_profile_map = {}

                for _, r_row in avg_df.iterrows():
                    pid_key = int(r_row["Player ID"])

                    history_profile_map[pid_key] = {
                        "GR": float(r_row.get("Guess rate",     0.0)),
                        "UF": float(r_row.get("Usefulness",     0.0)),
                        "OP": float(r_row.get("OP guess rate",  0.0)),
                        "ED": float(r_row.get("ED guess rate",  0.0)),
                        "IN": float(r_row.get("IN guess rate",  0.0))
                    }
                
                alias_txt_path      = self.tour_dir / FILE_ALIAS
                current_alias_lines = []

                if alias_txt_path.exists():
                    with open(alias_txt_path, "r", encoding = "utf-8") as f_alias:
                        for a_line in f_alias:
                            if "," in a_line: current_alias_lines.append(a_line.strip().split(","))
                
                id_table_lookup = {r[0].strip().lower(): int(r[1]) for r in rows_ids[1:] if len(r) >= 2 and r[0] and r[1]}
                
                with open(alias_txt_path, "w", encoding = "utf-8") as f_out:
                    for parts in current_alias_lines:
                        p_name      = parts[0].strip()
                        alias_name  = parts[1].strip()
                        p_low       = p_name.lower()
                        
                        p_id = id_table_lookup.get(p_low)
                        if p_id is None and alias_name.lower() in id_table_lookup: p_id = id_table_lookup.get(alias_name.lower())
                            
                        if p_id in history_profile_map:
                            h_prof = history_profile_map[p_id]
                            f_out.write(f"{p_name}, {alias_name}, {h_prof['GR']:.2f}, {h_prof['UF']:.2f}, {h_prof['OP']:.2f}, {h_prof['ED']:.2f}, {h_prof['IN']:.2f}\n")

                        else: f_out.write(f"{p_name}, {alias_name}, N/A, N/A, N/A, N/A, N/A\n")

                print("[✓] Historic baselines saved to alias.txt")

            except Exception as e:
                print(f"[!] Failed to fetch historic baselines: {e}")
                print("[?] Continuing structural pipeline execution, ignoring baseline fetching")

        dry_mode_mapping    = {
            "Usual"                 : "1",
            "Watched"               : "2",
            "Watched OP"            : "3",
            "Watched ED"            : "4",
            "Watched IN"            : "5",
            "Watched IN -Chanting"  : "6",
            "Watched OPED"          : "7",
            "Watched 2+8s"          : "8",
            "Watched 5s"            : "9",
            "Watched -2009"         : "10",
            "Random OP"             : "11",
            "Random ED"             : "12",
            "Random IN"             : "13",
            "Random OPED"           : "14",
            "Random Chanting"       : "15",
            "Other Random"          : "16",
            "Other Watched"         : "17",
            "Brute-force"           : "18",
        }

        if "sub_results" in meta_res:
            for sub_player, replaced_player in meta_res["sub_results"].items():
                s_low                       = sub_player.lower()
                chosen_team_id, chosen_tier = self.assignments[replaced_player.lower()]
                self.assignments[s_low]     = (chosen_team_id, chosen_tier)

                self.rosters[chosen_team_id].add(sub_player)

                self.subbed_players_set.add(s_low)
                self.subbed_players_set.add(replaced_player.lower())

                self.sub_relations[replaced_player.casefold()].append(sub_player)
                self.sub_relations[s_low] = [replaced_player]

        if not self.tour_label: self.tour_label = init_label

        self.val_str            = meta_res["th_str"]
        self.base_exp           = meta_res["base_exp"]
        self.new_players        = meta_res["selected_new"]
        self.dry_choice         = meta_res.get("dry_choice", "No")
        self.share_choice       = meta_res.get("share_choice", "No, I'll upload the site folder to Netlify Drop post-tour")
        self.exp_map            = {name: (self.base_exp - 1 if name.lower() in self.subbed_players_set else self.base_exp) for name in all_known}
        self.dry_mapped_index   = dry_mode_mapping.get(self.tour_label, "15")

        return True

    def process_and_generate(self):
        watched_valid = self.missing_list_count <= THRESH_WTCH

        for path in self.json_paths:
            with open(path, encoding = "utf-8") as f: data = json.load(f)
            songs = data.get("songs", [])

            if not path.stem.startswith("amq"):
                match_digits = re.search(r'(\d+)$', path.stem)

                if match_digits:
                    m = int(match_digits.group(1))
                    if m <= THRESH_SONG: songs = songs[:min(m, len(songs))]

            if not songs: continue

            raw_f_players = set()

            for s in songs:
                for p in s.get("correctGuessPlayers", []):
                    if      isinstance(p, str)                  : raw_f_players.add(p)
                    elif    isinstance(p, dict) and "name" in p : raw_f_players.add(p["name"])

                for ls in s.get("listStates", []):
                    if "name" in ls: raw_f_players.add(ls["name"])

            final_members = set(raw_f_players)

            if self.use_teams:
                t_in_f = {self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments}

                for tid in t_in_f:
                    ros     = self.rosters[tid]

                    for m_p in ros:
                        if m_p.lower() not in self.assignments:
                            for c_p in raw_f_players:
                                if c_p.lower() in self.assignments and self.assignments[c_p.lower()][0] == tid:
                                    self.assignments[m_p.lower()] = self.assignments[c_p.lower()]

                if len(final_members) < 8:
                    for tid in t_in_f: final_members.update(self.rosters[tid])

            apply_rev       = (len(final_members) % 2 == 0)
            max_s           = max(s.get("songNumber", 0) for s in songs)
            f_type_totals   = defaultdict(int)

            for song in songs:
                st = song.get("songInfo", {}).get("type")
                if st in [1, 2, 3]: f_type_totals[st] += 1

            for name in final_members:
                if name in raw_f_players:
                    self.s_part[name] += max_s
                    for t in [1, 2, 3]: self.p_type_s[name][t] += f_type_totals[t]

            for song in songs:
                si = song.get("songInfo", {})

                st      = si.get("type",        3)
                t_num   = si.get("typeNumber",  0)

                ann_id  = str(si.get("annSongId"))
                is_chan = ann_id in self.chanting_ids

                anime_name  = si.get("animeNames",  {})         .get("romaji", "Unknown")
                song_name   = si.get("songName",    "Unknown")
                artist_name = si.get("artist",      "Unknown")

                if len(anime_name)  > THRESH_CHRL: anime_name   = re.sub(r'\s+\S*$', '', anime_name     [:THRESH_CHRL]) + " ..."
                if len(song_name)   > THRESH_CHRM: song_name    = re.sub(r'\s+\S*$', '', song_name      [:THRESH_CHRM]) + " ..."
                if len(artist_name) > THRESH_CHRS: artist_name  = re.sub(r'\s+\S*$', '', artist_name    [:THRESH_CHRS]) + " ..."

                if      st == 1 : type_fmt = f"(OP{t_num})"
                elif    st == 2 : type_fmt = f"(ED{t_num})"
                else            : type_fmt = f"(IN)"

                song_line   = f"{anime_name} {type_fmt}: {song_name} by {artist_name}"
                raw_correct = song.get("correctGuessPlayers", [])
                correct     = set()

                for p in raw_correct:
                    if      isinstance(p, str)                  : correct.add(p)
                    elif    isinstance(p, dict) and "name" in p : correct.add(p["name"])

                active_correct  = correct & final_members
                amt_correct     = len(active_correct)

                self.song_history.append((correct, raw_f_players))

                ls                          =   song.get("listStates", [])
                self.global_stats["tot_c"]  +=  len(correct)

                try:
                    vint_raw = si.get("vintage", "")
                    yr          = int   (extract_year   (vint_raw))     if vint_raw else None
                    vint_scaled = float (extract_year   (vint_raw))     if vint_raw else 0.0
                    vint_pretty = format_year           (vint_scaled)   if vint_raw else "Unknown"

                except: 
                    yr          = None
                    vint_pretty = "Unknown"

                if yr is not None: self.all_vint.append(yr)

                raw_diff = si.get("animeDifficulty")

                try     : safe_diff = float(raw_diff)
                except  : safe_diff = 0.0

                self.all_diff.append(safe_diff)

                self.song_data.append({
                    "vintage"       : yr if yr is not None else 0,
                    "difficulty"    : safe_diff,
                    "correct_count" : int(len(correct))
                })

                song_line_hover = f"{anime_name} {type_fmt}: {song_name} by {artist_name} ({vint_pretty}/{safe_diff:.2f}: {len(correct)}/8)"

                if isinstance(si.get("animeGenre"), list): self.genre_c .update(si.get("animeGenre"))
                if isinstance(si.get("animeTags"),  list): self.tag_c   .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])

                if yr is not None and yr > 0:
                    diffs_arr   = [s["difficulty"] for s in self.song_data]
                    max_diff_v  = max(diffs_arr) if diffs_arr else 0
                    num_x_v     = 8 if max_diff_v < 40 else 9
                    num_y_v     = 8 if max_diff_v < 40 else 9
                    x_idx_v     = min(int(math.floor(safe_diff / 5)), num_x_v - 1)
                    vint_floor  = math.floor(float(yr))

                    if num_y_v == 8 : y_idx_v = 0 if vint_floor < 1995 else min(int(math.floor((vint_floor - 1995) / 5)) + 1, 7)
                    else            : y_idx_v = 0 if vint_floor < 1990 else min(int(math.floor((vint_floor - 1990) / 5)) + 1, 8)

                    self.matrix_song_details[f"{x_idx_v}-{y_idx_v}"].append(song_line_hover)

                if len(correct) == 0: self.tour_song_details["Total 0/8s"].append(song_line)

                elif len(correct) == 1:
                    sw_v = list(correct)[0]
                    self.tour_song_details["Total 1/8s"].append(f"{song_line} ({sw_v})")
                    if sw_v.lower() in self.assignments:
                        self.team_song_details[self.assignments[sw_v.lower()][0]]["Total 1/8s"].append(f"{song_line} ({sw_v})")

                elif len(correct) == 2:
                    p_list_v = list(correct)
                    self.tour_song_details["Total 2/8s"].append(f"{song_line} ({p_list_v[0]}/{p_list_v[1]})")

                elif apply_rev and len(final_members - correct) == 1:
                    missing_player_v = list(final_members - correct)[0]
                    self.tour_song_details["Total 7/8s"].append(f"{song_line} ({missing_player_v})")

                elif len(final_members - correct) == 0: self.tour_song_details["Total 8/8s"].append(song_line)

                for sw_v in active_correct:
                    if amt_correct == 1: self.player_song_details[sw_v]["1/8s"].append(song_line)

                    elif amt_correct == 2:
                        opp_player_v = list(active_correct)[1]                                  if sw_v.casefold() == list(active_correct)[0].casefold() and len(active_correct) > 1    else list(active_correct)[0]
                        t_sw_v       = self.assignments.get(sw_v.lower          (), (None,))[0] if self.use_teams                                                                       else None
                        t_opp_v      = self.assignments.get(opp_player_v.lower  (), (None,))[0] if self.use_teams                                                                       else None

                        if t_sw_v is not None and t_opp_v is not None and t_sw_v == t_opp_v : self.player_song_details[sw_v]["2/8s"].append(f"{song_line} (covered by {opp_player_v})")
                        else                                                                : self.player_song_details[sw_v]["2/8s"].append(f"{song_line} (blocked by {opp_player_v})")

                    def extract_year_float(vint_str):
                        if not vint_str: return 0.0
                        vint_str = vint_str.strip()

                        match = re.search(r'(\d{4})', vint_str)
                        if not match: return 0.0

                        year        = float(match.group(1))
                        low_str     = vint_str.lower()

                        if      "winter"    in low_str: return year + 0.00
                        elif    "spring"    in low_str: return year + 0.25
                        elif    "summer"    in low_str: return year + 0.50
                        elif    "fall"      in low_str: return year + 0.75

                        return year

                    if safe_diff > 0    : self.p_hit_diff[sw_v].append(safe_diff)
                    if yr is not None   : self.p_hit_vint[sw_v].append(extract_year_float(si.get("vintage")))

                if apply_rev and len(final_members - correct) == 1:
                    missing_player_v = list(final_members - correct)[0]
                    self.player_song_details[missing_player_v]["7/8s"].append(song_line)

                if isinstance(si.get("animeGenre"), list):
                    for gen in si.get("animeGenre"): self.tour_song_details[f"Genre: {gen}"].append(song_line)

                if isinstance(si.get("animeTags"), list):
                    for tag in si.get("animeTags"):
                        if tag not in EXCLUDED_TAGS: self.tour_song_details[f"Tag: {tag}"].append(song_line)

                if si.get("vintage"):
                    for p in song.get("correctGuessPlayers", []):
                        p_name_v = p if isinstance(p, str) else p.get("name") if isinstance(p, dict) else None
                        if p_name_v: self.raw_vintage_by_guess[p_name_v].append(si.get("vintage"))

                    for ls_v in song.get("listStates", []):
                        if "name" in ls_v: self.raw_vintage_by_list[ls_v["name"]].append(si.get("vintage"))

                seen_song_times = set()

                if isinstance(raw_correct, list):
                    for p in raw_correct:
                        if isinstance(p, dict) and "name" in p and "answerTime" in p:
                            try     : seen_song_times.add((str(p["name"]).casefold(), float(p["answerTime"])))
                            except  : pass

                for key_name in ["answerTimes", "answerTime", "answerTimesByPlayer", "playerAnswerTimes"]:
                    val = song.get(key_name)

                    if isinstance(val, dict):
                        for p_name, t_val in val.items():
                            try                             : seen_song_times.add((str(p_name).casefold(), float(t_val)))
                            except (ValueError, TypeError)  : pass

                    elif isinstance(val, list):
                        for item in val:
                            if isinstance(item, dict):
                                p_name  = item.get("name") or item.get("player")        or item.get("playerName")
                                t_val   = item.get("time") or item.get("answerTime")    or item.get("value")

                                if p_name and t_val is not None:
                                    try                             : seen_song_times.add((str(p_name).casefold(), float(t_val)))
                                    except (ValueError, TypeError)  : pass

                name_map = {m.lower(): m for m in final_members}

                for p_name_lower, t_float in seen_song_times:
                    if p_name_lower in name_map: self.p_answer_times[name_map[p_name_lower]].append(t_float)

                s_riggers = {p["name"] for p in ls}

                if len(ls) == 1:
                    u                   =   ls[0]["name"]
                    self.p_l_solos[u]   +=  1
                    if not (len(correct) == 1 and list(correct)[0] == u): self.p_m_erigs[u] += 1

                if ls:
                    is_true_solo_rig = (len(ls) == 1)

                    for p in ls:
                        n_v       = p["name"]
                        marker_v  = "✓" if (n_v in active_correct) else "✗"

                        self.player_song_details[n_v]["Rigs"].append(f"{marker_v} {song_line}")

                        if is_true_solo_rig:
                            self.player_song_details[n_v]["Solo Rigs"].append(f"{marker_v} {song_line}")

                            s_v     = sorted(list(active_correct - {n_v}))
                            sC_v    = len(s_v)

                            if      sC_v == 0   : tag_v = "(0/8)"
                            elif    sC_v == 1   : tag_v = f"(stolen by {s_v[0]})"
                            elif    sC_v == 2   : tag_v = f"(stolen by {s_v[0]}/{s_v[1]})"
                            else                : tag_v = f"({amt_correct}/8)"

                            if n_v in active_correct and amt_correct == 1   : self.player_song_details[n_v]["Solo Rig Conversions"].append(f"✓ {song_line}")
                            else                                            : self.player_song_details[n_v]["Solo Rig Conversions"].append(f"✗ {song_line} {tag_v}")

                lives_taken_players = []
                lives_saved_players = []

                if self.use_teams:
                    t_list = list({self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments})

                    if len(t_list) == 2:
                        tA, tB = t_list[0],             t_list[1]
                        cA, cB = correct & self.rosters[tA], correct & self.rosters[tB]

                        if len(cA) == 4 and not cB: self.t_sweeps[tA] += 1; self.global_stats["sweeps"] += 1
                        if len(cB) == 4 and not cA: self.t_sweeps[tB] += 1; self.global_stats["sweeps"] += 1

                        if (len(cA & final_members) == 4 and not (cB & final_members)) or (len(cB & final_members) == 4 and not (cA & final_members)):
                            self.tour_song_details["Total 4-0s"].append(song_line)

                        for cur, opp in [(tA, tB), (tB, tA)]:
                            cC, oC = correct & self.rosters[cur], correct & self.rosters[opp]

                            if not oC: 
                                for p in cC: 
                                    self.p_pts[p] += 1
                                    lives_taken_players.append(p.lower())

                            if len(cC) == 1 and len(oC) > 0: 
                                lone_p = list(cC)[0]
                                self.p_blks[lone_p] += 1
                                lives_saved_players.append(lone_p.lower())

                        for _, opp_v, cC_v, oC_v in [(tA, tB, cA & final_members, cB & final_members), (tB, tA, cB & final_members, cA & final_members)]:
                            oL_v = self.t1_lookup.get(opp_v, f"Team {opp_v}")

                            if not oC_v:
                                for p_v in cC_v: self.player_song_details[p_v]["Lives Taken"].append(f"{song_line} (from Team {oL_v})")

                            if len(cC_v) == 1 and len(oC_v) > 0:
                                oP_v = sorted(list(oC_v), key = lambda x: self.assignments.get(x.lower(), (None, "5"))[1])

                                if      len(oP_v) == 1  : opp_tag_v = f"(from {oP_v[0]} in Team {oL_v})"          if oP_v[0] != oL_v                  else f"(from {oP_v[0]})"
                                elif    len(oP_v) == 2  : opp_tag_v = f"(from {oP_v[0]}/{oP_v[1]} in Team {oL_v})"  if oP_v[0] != oL_v and oP_v[1] != oL_v  else f"(from {oP_v[0]}/{oP_v[1]})"
                                else                    : opp_tag_v = f"(from Team {oL_v})"

                                self.player_song_details[list(cC_v)[0]]["Lives Saved"].append(f"{song_line} {opp_tag_v}")

                    for tid in t_list:
                        ros     = self.rosters[tid]
                        c_on_t  = correct & ros

                        self.t_c_ps[tid].append(len(c_on_t) / 4.0)
                        if yr is not None: self.t_vint[tid].append(yr)

                        if s_riggers & ros:
                            self.t_on_syn[tid].append(len(c_on_t)                   / 4.0)
                            self.t_sh_rig[tid].append((len(s_riggers & ros) - 1)    / 3.0)

                        else: self.t_off_syn[tid].append(len(c_on_t) / 4.0)

                if len(final_members - correct) == 0: self.global_stats["fulls"] += 1

                elif apply_rev and len(final_members - correct) == 1:
                    self.global_stats   ["sevens"]                          += 1
                    self.p_rev_e        [list(final_members - correct)[0]]  += 1

                elif len(correct) == 2:
                    self.global_stats["doubles"] += 1

                    p_list = list(correct)
                    p1, p2 = p_list[0], p_list[1]
                    
                    self.p_two_e[p1] += 1
                    self.p_two_e[p2] += 1

                    t1 = self.assignments.get(p1.lower(), (None,))[0] if self.use_teams else None
                    t2 = self.assignments.get(p2.lower(), (None,))[0] if self.use_teams else None

                    if t1 is not None and t2 is not None and t1 == t2:
                        rel_p1_by_p2    = f" (covered by {p2})"
                        rel_p2_by_p1    = f" (covered by {p1})"
                        rel_both        = f" ({p1} and {p2})"

                    else:
                        rel_p1_by_p2    = f" (blocked by {p2})"
                        rel_p2_by_p1    = f" (blocked by {p1})"
                        rel_both        = f" ({p1} and {p2})"

                    song["blocked_p1_by_p2"]    = rel_p1_by_p2
                    song["blocked_p2_by_p1"]    = rel_p2_by_p1
                    song["both_players"]        = rel_both

                elif len(correct) == 1:
                    self.global_stats["solos"]  +=  1
                    sw                          =   list(correct)[0]
                    self.e_counts[sw]           +=  1

                    if sw.lower() in self.assignments: self.t_solos[self.assignments[sw.lower()][0]] += 1

                elif len(correct) == 0: self.global_stats["blanks"] += 1

                amtcorrect = len(correct)

                if amtcorrect > 0:
                    teamsize    = 4
                    uf_song     = 0.0

                    for i in range(teamsize): uf_song += math.comb(2 * teamsize - (i + 2), amtcorrect - 1) / math.comb(2 * teamsize - 1, amtcorrect - 1)
                    uf_song /= teamsize
                    
                    for name in final_members:
                        if name in correct: self.p_usefulness_sum[name] += uf_song

                for name in final_members:
                    if name in correct:
                        self.c_counts       [name] += 1
                        self.p_overs_sum    [name] += len(correct)

                        if st in [1, 2, 3]: 
                            self.p_type_c[name][st] += 1
                            self.player_song_details[name][f"Type {st}"].append(f"✓ {song_line}")

                        if is_chan:          
                            self.p_chan_c[name] += 1
                            self.player_song_details[name]["Chant"].append(f"✓ {song_line}")

                        if yr is not None: self.p_c_vint[name].append(yr)
                        self.player_song_details[name]["Overall"].append(f"✓ {song_line}")

                    else:
                        if st in [1, 2, 3]  : self.player_song_details[name][f"Type {st}"]  .append(f"✗ {song_line}")
                        if is_chan          : self.player_song_details[name]["Chant"]       .append(f"✗ {song_line}")

                        self.player_song_details[name]["Overall"].append(f"✗ {song_line}")

                    if is_chan: self.p_chan_s[name] += 1

                if ls:
                    for p in ls:
                        n = p["name"]
                        self.p_rigs[n] += 1

                        if n in correct: 
                            self.p_rigs_h[n] += 1
                            self.p_lh_corr[n].append(len(correct))
                            if yr is not None: self.p_lh_vint[n].append(yr)

                        if yr is not None: self.p_l_vint[n].append(yr)
                        self.p_l_corr[n].append(len(correct))

        if "Eru" in self.tour_label and self.use_teams:
            self.p_pts  .clear()
            self.p_blks .clear()

            for cor, raw_f_players in self.song_history:
                t_list = list({self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments})

                if len(t_list) == 2:
                    tA, tB  = t_list[0], t_list[1]
                    cA      = {self.assignments[p.lower()][1]: p for p in raw_f_players if p.lower() in self.assignments and self.assignments[p.lower()][0] == tA}
                    cB      = {self.assignments[p.lower()][1]: p for p in raw_f_players if p.lower() in self.assignments and self.assignments[p.lower()][0] == tB}

                    for tr in ["1", "2", "3", "4"]:
                        pA, pB = cA.get(tr), cB.get(tr)

                        if pA and pB:
                            rA, rB = pA in cor, pB in cor

                            if rA and not rB: self.p_pts[pA] += 1
                            if rB and not rA: self.p_pts[pB] += 1
                            if rA and rB:
                                self.p_blks[pA] += 0.50
                                self.p_blks[pB] += 0.50

        final_threshold = 6 if len(self.s_part) <= THRESH_PLYR else 5

        if      self.base_exp >= final_threshold: stage = "Final"
        elif    self.base_exp == 3              : stage = "Mid-Tour"
        else                                    : stage = f"R{self.base_exp}"

        prefix      = f"{self.tour_label.strip()} Tour, " 
        png_path    = self.tour_dir / "png"
        web_path    = self.tour_dir / "site"

        for path in [png_path, web_path]:
            path.mkdir(parents = True, exist_ok = True)

            for item in path.iterdir():
                if item.is_file(): item.unlink()

        tasks = []

        tasks.append((self._create_player_png,      (self.elo_map, watched_valid, stage, png_path, self.apps, prefix, self.exp_map, self.base_exp, self.new_players, self.val_str)))
        tasks.append((self._create_tour_png,        (self.use_teams, watched_valid, png_path)))
        tasks.append((self._create_scatter_png,     (png_path, False, self.elo_map)))
        tasks.append((self._create_song_png,        (png_path, )))
        tasks.append((self._create_dashboard_html,  (web_path, self.use_teams, watched_valid)))

        if self.assignments:
            tasks.append((self._create_tier_png, (self.assignments, png_path, any(self.p_chan_s.values()))))
            tasks.append((self._create_team_png, (self.assignments, self.t1_lookup, png_path)))

        if watched_valid: tasks.append((self._create_scatter_png, (png_path, True, self.elo_map)))

        with fut.ProcessPoolExecutor() as executor:
            task = {executor.submit(func, *args): func.__name__ for func, args in tasks}

            for future in fut.as_completed(task):
                task_name = task[future]

                try                     : future.result()
                except Exception as e   : print(f"Task {task_name} failed: {e}")

        self._fuse(png_path)
        allowed_files = {"General.png", "Player.png", "Extra.png", "Plots.png"}

        for file_path in png_path.glob("*.png"):
            if file_path.name not in allowed_files:
                try                 : file_path.unlink()
                except Exception    : pass

        if self.dry_choice != "No":            
            target_jsons_dir    = self.script_dir.parent    / "jsons"
            target_codes_file   = self.script_dir.parent    / "codes.txt"
            source_jsons_dir    = self.tour_dir             / "json"
            source_codes_file   = self.tour_dir             / "code.txt"
            script_name         = "ngm_local.py" if "don't push" in self.dry_choice else "ngm_stats.py"
            target_script       = self.script_dir.parent    / script_name

            print(f"[?] Processing Tour {self.tour_id} using Dry's script")

            for file in self.script_dir     .glob("*.png")  : file.unlink()
            for file in target_jsons_dir    .glob("*.json") : file.unlink()

            if source_jsons_dir.exists():
                for file in source_jsons_dir.glob("*.json"): shutil.copy(file, target_jsons_dir / file.name)

            print("[✓] Copied JSONs to Dry's workspace")

            if source_codes_file.exists():
                shutil.copy(source_codes_file, target_codes_file)
                print("[✓] Copied code.txt to Dry's workspace")

            print(f"[?] Running Dry's script")

            try:
                subprocess.run([sys.executable, str(target_script)], cwd = str(self.script_dir.parent), check = True)
                print(f"[✓] Ran Dry's script successfully")

                output_dir = self.tour_dir / "dry"
                output_dir.mkdir(parents = True, exist_ok = True)

                files_to_copy = {
                    "Stats.png"                         : "1-Player.png",
                    "Stats2.png"                        : "2-Type.png",
                    "Stats3 - Watched Exclusive.png"    : "3-List.png",
                    "Stats Songs.png"                   : "4-Song.png",
                    "Stats4.png"                        : "5-Extra.png"
                }

                print("[?] Copying Dry's PNGs back")

                for src_name, dest_name in files_to_copy.items():
                    src_file    = self.script_dir.parent    / src_name
                    dest_file   = output_dir                / dest_name

                    if src_file.exists():
                        shutil.copy(src_file, dest_file)
                        print(f"[✓] Copied {src_name} as {dest_name}")

                    else: print(f"[X] {src_name} not found in Dry's workspace")

            except subprocess.CalledProcessError as e: (f"[X] Failed to run Dry's script: {e}")

        if self.share_choice == "No, I'll upload the site folder to Netlify Drop post-tour":
            workspace_root  = self.script_dir.parent
            png_src         = self.tour_dir / "png"
            site_src        = self.tour_dir / "site"
            player_file     = png_src / "Player.png"
            extra_file      = png_src / "Extra.png"

            if player_file  .exists(): shutil.copy      (player_file,   workspace_root / f"hako-{self.tour_id}-player.png")
            if extra_file   .exists(): shutil.copy      (extra_file,    workspace_root / f"hako-{self.tour_id}-extra.png")
            if site_src     .exists(): shutil.copytree  (site_src,      workspace_root / f"hako-{self.tour_id}-upload",     dirs_exist_ok = True)

        elif self.share_choice == "Yes, push it to the archive":
            timestamp   = datetime.datetime.now().strftime("%y%m%d%H%M")
            archive_dir = self.script_dir.parent / "hako" / "archive" / timestamp
            hako_dir    = web_path 

            print(f"[?] Copying files to archive/{timestamp}")

            try:
                shutil.copytree(hako_dir, archive_dir, dirs_exist_ok = True)
                print(f"[✓] Copied all files from {hako_dir.name}")

            except Exception as e: print(f"[X] Failed to copy files: {e}")
  
            dashboard_url = f"https://frittutisna.github.io/Stats-Maker/hako/archive/{timestamp}/index.html?update=1"
            print(f"[?] Pushing to GitHub")

            try:
                subprocess.run(["git", "add", "."],                     check = True)
                subprocess.run(["git", "commit", "-m", "Updated tour"], check = True)
                subprocess.run(["git", "push"],                         check = True)

                print(f"[✓] Deployment completed, dashboard link: {dashboard_url}")

            except subprocess.CalledProcessError as git_error: print(f"[X] Failed to push to GitHub: {git_error}")

    def _scan_players(self, paths):
        players = set           ()
        apps    = defaultdict   (set)

        for p in paths:
            try:
                with open(p, encoding = "utf-8") as f:
                    data = json.load(f)

                    for s in data.get("songs", []):
                        for plyr in s.get("correctGuessPlayers", []):
                            if isinstance(plyr, str):
                                players.add(plyr)
                                apps[plyr].add(str(p))

                            elif isinstance(plyr, dict) and "name" in plyr:
                                players.add(plyr["name"])
                                apps[plyr["name"]].add(str(p))

                        for ls in s.get("listStates", []): players.add(ls["name"])
                        apps[ls["name"]].add(str(p))

            except: continue

        return players, apps

    def _load_team_data(self, all_known):
        codes = self.tour_dir / FILE_CODES
        if not codes.exists() or os.path.getsize(codes) == 0: return False, {}, {}, {}, defaultdict(set), all_known, [], list(all_known)
        with open(codes, "r", encoding = "utf-8") as f: lines = [line.strip() for line in f if line.strip()]

        has_avg     = any(l.lower().startswith(("average", "avg")) for l in lines)
        team_lines  = 0
        bad_lines   = []

        for line in lines:
            if line.lower().startswith(("average", "avg", "sub")) or line.startswith("http"): continue
            team_text = line.split("|", 1)[0].strip()

            if not re.findall(TEAMS_RE, team_text)  : bad_lines.append(line)
            else                                    : team_lines += 1

        if bad_lines or team_lines == 0 or not has_avg:
            error_details = ["Broken codes.txt"]

            if bad_lines    : error_details.append(f"Cannot parse line structure: '{bad_lines[0]}'")
            if not has_avg  : error_details.append("Missing (Average: Elo) line")

            messagebox.showerror("Invalid Code File", "\n".join(error_details))
            return False

        self.main_roster_names                      = set()
        self.sub_relations                          = defaultdict(list)
        elo_map, assignments, rosters, t1_lookup    = {}, {}, defaultdict(set), {}
        avail                                       = sorted(list(all_known))
        alias_path                                  = self.tour_dir / FILE_ALIAS
        local_aliases                               = {}

        if alias_path.exists():
            with open(alias_path, "r", encoding = "utf-8") as f:
                for line in f:
                    if "," in line:
                        k, v = line.strip().split(",", 1)
                        local_aliases[k.strip().lower()] = v.strip()

        new_aliases = {}

        def find_best_match(p_in):
            p_low = p_in.lower()

            if p_low in local_aliases:
                m = local_aliases[p_low]
                if m in all_known: return m

            match = next((n for n in all_known if n.lower() == p_low), None)

            if not match:
                if not self.id_database: self.id_database = self._load_player_ids()

                if p_low in self.id_database:
                    target_id   = self.id_database[p_low]
                    match       = next((n for n in all_known if self.id_database.get(n.lower()) == target_id), None)

            if match: new_aliases[p_in] = match
            return match

        for line in lines:
            matches = re.findall(TEAMS_RE, line)

            for p_in, val in matches:
                if not line.lower().startswith("subs:"):
                    match = find_best_match(p_in)
                    if match: elo_map[match.lower()] = val

        idx                 = 1
        sub_candidates_raw  = []

        for line in lines:
            if line.lower().startswith("subs:"):
                mems_subs = re.findall(TEAMS_RE, line)

                for p_sub, val_s in mems_subs:
                    m_sub = find_best_match(p_sub)

                    if m_sub:
                        self.subbed_players_set.add(m_sub.lower())
                        sub_candidates_raw.append(m_sub)
                        elo_map[m_sub.lower()] = val_s

                continue

            mems = re.findall(TEAMS_RE, line.split("|")[0])
            if not mems: continue

            p_captain, _    = mems[0]
            c_match         = find_best_match(p_captain)
            ename           = c_match if c_match else p_captain
            t1_lookup[idx]  = ename

            for i, (p_in, _) in enumerate(mems[:4]):
                tier    = str(i + 1)
                match   = find_best_match(p_in)

                if match:
                    self.main_roster_names.add(match.lower())
                    assignments[match.lower()] = (idx, tier)
                    rosters[idx].add(match)
                    if match in avail: avail.remove(match)

            idx += 1

        original_players_display = [p for p in all_known if p.lower() in assignments and p.lower() not in [s.lower() for s in sub_candidates_raw]]

        if new_aliases:
            existing_pairs = set()

            if alias_path.exists():
                with open(alias_path, "r", encoding = "utf-8") as f:
                    for l in f:
                        parts = [p.strip().lower() for p in l.split(",")]
                        if len(parts) >= 2: existing_pairs.add((parts[0], parts[1]))

            with open(alias_path, "a", encoding = "utf-8") as f:
                for k, v in new_aliases.items():
                    k_low, v_low = k.strip().lower(), v.strip().lower()

                    if (k_low, v_low) not in existing_pairs:
                        f.write(f"{k}, {v}\n")
                        existing_pairs.add((k_low, v_low))

        return True, elo_map, assignments, t1_lookup, rosters, all_known, sub_candidates_raw, original_players_display

    def _compute_player_rows(self, elo_map, apps, exp_map, base_exp, new_players, watched, active, t_labels, avg_rank):
        rows, eligibility = [], []

        for name in self.s_part:
            tot, cor    = self.s_part[name], self.c_counts[name]
            target      = exp_map.get(name, base_exp)
            d_name      = name

            if name in new_players: d_name += " ★"

            if target != "ignore" and target < base_exp:
                if name.lower() in self.main_roster_names   : d_name += " ▼"
                else                                        : d_name += " ▲"

            is_eligible = not ("▼" in d_name or "▲" in d_name)
            eligibility.append(is_eligible)

            history_baselines   = {"GR": np.nan, "UF": np.nan, "OP": np.nan, "ED": np.nan, "IN": np.nan}
            alias_txt_path      = self.tour_dir / FILE_ALIAS

            if alias_txt_path.exists():
                try:
                    with open(alias_txt_path, "r", encoding = "utf-8") as f_alias:
                        for a_line in f_alias:
                            if "," in a_line:
                                p_splits = [x.strip() for x in a_line.split(",")]
                                if len(p_splits) >= 7 and (p_splits[0].lower() == name.lower() or p_splits[1].lower() == name.lower()):
                                    if p_splits[2] != "N/A":
                                        history_baselines = {
                                            "GR": float(p_splits[2]),
                                            "UF": float(p_splits[3]),
                                            "OP": float(p_splits[4]),
                                            "ED": float(p_splits[5]),
                                            "IN": float(p_splits[6])
                                        }

                                    break

                except Exception: pass

            row = {"Player": d_name}
            if self.use_teams: row["Elo"] = elo_map.get(name.lower(), np.nan)

            current_gr = cor / tot if tot else 0.0
            row.update({"GR": current_gr})

            delta_gr = (current_gr * 100) - history_baselines["GR"] if pd.notnull(history_baselines["GR"])  else np.nan
            row.update({"GR Δ": round(delta_gr, 2)                   if pd.notnull(delta_gr)                 else np.nan})

            if self.use_teams: 
                uf_val = (self.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0
                row.update({"UF": uf_val})

                delta_uf = uf_val - history_baselines["UF"] if pd.notnull(history_baselines["UF"])  else np.nan
                row.update({"UF Δ": round(delta_uf, 2)       if pd.notnull(delta_uf)                 else np.nan})

                try     : elo = float(elo_map.get(name.lower(), 0.0))
                except  : elo = 0.0

                all_ufs     = [(self.p_usefulness_sum[p] * avg_rank * 8) / self.s_part[p] if self.s_part[p] else 0.0 for p in self.s_part]
                all_elos    = []

                for p in self.s_part:
                    try     : all_elos.append(float(elo_map.get(p.lower(), 0.0)))
                    except  : all_elos.append(0.0)

                if len(all_elos) > 1 and np.var(all_elos) > 0:
                    slope, intercept    = np.polyfit(all_elos, all_ufs, 1)
                    res_std             = np.std(np.array(all_ufs) - (slope * np.array(all_elos) + intercept))

                    if res_std == 0: res_std = 1

                else: slope, intercept, res_std = 0, np.mean(all_ufs) if all_ufs else 0, 1

                residual    = uf_val - (slope * elo + intercept)
                perf_score  = (1 / (1 + np.exp(SCALE_PERF * (residual / res_std)))) * 100

                row.update({"Score": perf_score})

            avg_over8 = self.p_overs_sum[name] / cor if cor else np.nan
            row.update({"1/8s": self.e_counts[name], "2/8s": self.p_two_e[name], "7/8s": self.p_rev_e[name], "Mean Over-8": avg_over8})
            if self.use_teams: row.update({"Lives Taken": self.p_pts[name], "Lives Saved": self.p_blks[name]})

            for tid in active:
                seen                = self.p_type_s[name][tid]
                current_type_gr     = self.p_type_c[name][tid] / seen if seen else np.nan
                row[t_labels[tid]]  = current_type_gr

                t_key           = t_labels[tid].split(" ")[0] 
                hist_base_val   = history_baselines.get(t_key, np.nan)
                
                if pd.notnull(current_type_gr) and pd.notnull(hist_base_val):
                    delta_type          = (current_type_gr * 100) - hist_base_val
                    row[f"{t_key} Δ"]   = round(delta_type, 2)

                else: row[f"{t_key} Δ"] = np.nan

            if watched:
                rig_over8 = np.mean(self.p_l_corr[name]) if self.p_l_corr[name] else np.nan

                row.update({
                    "Rigs"              : self.p_rigs[name],
                    "Rig Rate"          : self.p_rigs[name]             / tot                       if tot                          else np.nan,
                    "Solo Rigs"         : self.p_l_solos[name],
                    "Solo Rig Rate"     : self.p_l_solos[name]          / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Rig Over-8"        : rig_over8,
                    "Over-8 Δ"          : rig_over8 - avg_over8,
                    "Rig GR"            : self.p_rigs_h[name]           / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Off GR"            : (cor - self.p_rigs_h[name])   / (tot - self.p_rigs[name]) if (tot - self.p_rigs[name])    else np.nan,
                    "Rig Δ"             : (cor - self.p_rigs[name])     / cor                       if cor                          else np.nan,
                })

            h_diffs = self.p_hit_diff.get(name, [])
            h_vints = self.p_hit_vint.get(name, [])

            row.update({
                "Mean Difficulty Hit"   : np.mean   (h_diffs) if h_diffs else np.nan,
                "Median Vintage Hit"    : np.median (h_vints) if h_vints else np.nan
            })

            times       = self.p_answer_times.get(name, [])
            seen_chan   = self.p_chan_s[name]

            row["Median Time"]  = np.median(times)                  if times        else np.nan
            row["Chant GR"]     = self.p_chan_c[name] / seen_chan   if seen_chan    else np.nan

            rows.append(row)

        df = pd.DataFrame(rows)

        if "Score" in df.columns: df = df.sort_values(by = ["GR", "Score"], ascending = [False, False])

        elif "Elo" in df.columns:
            df["_sort_elo"] = pd.to_numeric(df["Elo"], errors = 'coerce')
            df              = df.sort_values(by = ["GR", "_sort_elo"], ascending = [False, True]).drop(columns = ["_sort_elo"])

        else: df = df.sort_values("GR", ascending = False)

        mask = pd.Series(eligibility, index = pd.DataFrame(rows).index).reindex(df.index).values
        return df, mask

    def _create_player_png(self, elo_map, watched, stage, path, apps, prefix, exp_map, base_exp, new_players, val_str):
        t_labels    = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
        active      = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]

        if len(active) <= 1 : active = []

        valid_elos  = [float(v) for v in elo_map.values() if str(v).replace('.', '', 1).isdigit() or (str(v).startswith('-') and str(v)[1:].replace('.', '', 1).isdigit())]
        avg_rank    = np.mean(valid_elos) if valid_elos else 1.0
        df, mask    = self._compute_player_rows(elo_map, apps, exp_map, base_exp, new_players, watched, active, t_labels, avg_rank)
        df_png      = df.copy()
        pcts        = ["GR"] + [t_labels[t] for t in active] + (["Rig Rate", "Solo Rig Rate", "Rig Δ", "Rig GR", "Off GR"] if watched else []) + ["Chant GR"]

        if "Elo"            in df_png.columns: df_png["Elo"]            = pd.to_numeric(df_png["Elo"],          errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "UF"             in df_png.columns: df_png["UF"]             = pd.to_numeric(df_png["UF"],           errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Score"          in df_png.columns: df_png["Score"]          = pd.to_numeric(df_png["Score"],        errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Median Time"    in df_png.columns: df_png["Median Time"]    = pd.to_numeric(df_png["Median Time"],  errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Mean Over-8"    in df_png.columns: df_png["Mean Over-8"]    = pd.to_numeric(df_png["Mean Over-8"],  errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Rig Over-8"     in df_png.columns: df_png["Rig Over-8"]     = pd.to_numeric(df_png["Rig Over-8"],   errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Over-8 Δ"       in df_png.columns: df_png["Over-8 Δ"]       = pd.to_numeric(df_png["Over-8 Δ"], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

        for c in pcts: df_png[c] = pd.to_numeric(df_png[c], errors = 'coerce').mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        delta_cols = ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]

        for dc in delta_cols:
            if dc in df_png.columns:
                df_png[dc] = pd.to_numeric(df_png[dc], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

        self._export_png(df_png, path, "Player.png", f"{prefix}Player Statistics, {stage}", mask, val_str)

    def _compute_tour_stats(self, use_teams, watched):
        def fmt_most(names, val):
            if not names: return "N/A"

            win = sorted(names, key = lambda x: (self.c_counts[x] / self.s_part[x]) if self.s_part[x] else 0)[0]
            gr  = (self.c_counts[win] / self.s_part[win]) * 100 if self.s_part[win] else 0

            return f"{win} ({val}{f', {gr:.2f}' if len(names) > 1 else ''})", win

        stats = [
            ["Median Vintage",  format_year(round(np.median(self.all_vint), 2))                         if self.all_vint    else "N/A", None],
            ["Mean Difficulty", f"{np.mean(self.all_diff):.2f}"                                         if self.all_diff    else "N/A", None],
            ["Mean GR",         f"{100 * (self.global_stats['tot_c'] / sum(self.s_part.values())):.2f}" if self.s_part      else "N/A", None],
            ["Total 0/8s",      self.global_stats["blanks"],    "Total 0/8s"],
            ["Total 1/8s",      self.global_stats["solos"],     "Total 1/8s"],
            ["Total 2/8s",      self.global_stats["doubles"],   "Total 2/8s"],
            ["Total 7/8s",      self.global_stats["sevens"],    "Total 7/8s"],
            ["Total 8/8s",      self.global_stats["fulls"],     "Total 8/8s"]
        ]

        if use_teams: stats.append(["Total 4-0s", self.global_stats["sweeps"], "Total 4-0s"])

        pop_gen = self.genre_c  .most_common(1)[0][0] if self.genre_c   else "N/A"
        pop_tag = self.tag_c    .most_common(1)[0][0] if self.tag_c     else "N/A"

        pop_gen_cnt = self.genre_c  .most_common(1)[0][1] if self.genre_c   else 0
        pop_tag_cnt = self.tag_c    .most_common(1)[0][1] if self.tag_c     else 0

        m1_p = [n for n, v in self.e_counts .items() if v == max(self.e_counts  .values(), default = 0) and v > 0]
        m2_p = [n for n, v in self.p_two_e  .items() if v == max(self.p_two_e   .values(), default = 0) and v > 0]
        m7_p = [n for n, v in self.p_rev_e  .items() if v == max(self.p_rev_e   .values(), default = 0) and v > 0]

        f1, w1 = fmt_most(m1_p, max(self.e_counts   .values(), default = 0))        
        f2, w2 = fmt_most(m2_p, max(self.p_two_e    .values(), default = 0))
        f7, w7 = fmt_most(m7_p, max(self.p_rev_e    .values(), default = 0))

        stats.extend([
            ["Most Popular Genre",  f"{pop_gen} ({pop_gen_cnt})" if self.genre_c    else "N/A", f"Genre: {pop_gen}"],
            ["Most Popular Tag",    f"{pop_tag} ({pop_tag_cnt})" if self.tag_c      else "N/A", f"Tag: {pop_tag}"],
            ["Most 1/8s",           f1, ("1/8s", w1)],
            ["Most 2/8s",           f2, ("2/8s", w2)],
            ["Most 7/8s",           f7, ("7/8s", w7)]
        ])

        plist   = list(self.s_part.keys())
        no_s    = sorted([n for n in plist if self.e_counts[n] ==   0 and self.s_part[n] > 0], key = lambda x: self.c_counts[x] / self.s_part[x], reverse = True)
        yes_s   = sorted([n for n in plist if self.e_counts[n] >    0 and self.s_part[n] > 0], key = lambda x: self.c_counts[x] / self.s_part[x])

        if no_s     : stats.append(["Highest GR Without 1/8s",  f"{no_s     [0]} ({100 * (self.c_counts[no_s    [0]] / self.s_part[no_s     [0]]):.2f})", None])
        if yes_s    : stats.append(["Lowest GR With 1/8s",      f"{yes_s    [0]} ({100 * (self.c_counts[yes_s   [0]] / self.s_part[yes_s    [0]]):.2f}, {self.e_counts[yes_s[0]]})", ("1/8s", yes_s[0])])

        if watched:
            conv        = []
            eligible    = [p for p in plist if self.p_l_solos[p] > 0]

            if eligible:
                total_hits      = sum((self.p_l_solos[p] - self.p_m_erigs[p]) for p in eligible)
                total_attempts  = sum(self.p_l_solos[p] for p in eligible)
                global_avg      = total_hits / total_attempts if total_attempts > 0 else 0

                for n in eligible:
                    t               = self.p_l_solos[n]
                    h               = t - self.p_m_erigs[n]
                    weighted_score  = (h + CONST_CONV * global_avg) / (t + CONST_CONV)
                    conv.append({'n': n, 'score': weighted_score, 'p': 100 * h / t, 'h': h, 't': t})

                b = sorted(conv, key = lambda x: x['score'], reverse = True)    [0]
                w = sorted(conv, key = lambda x: x['score'])                    [0]

                stats.append(["Best Solo Rig Converter",    f"{b['n']} ({b['p']:.2f}, {b['h']}/{b['t']})", ("Solo Rigs", b['n'])])
                stats.append(["Worst Solo Rig Converter",   f"{w['n']} ({w['p']:.2f}, {w['h']}/{w['t']})", ("Solo Rigs", w['n'])])

        return stats

    def _create_tour_png(self, use_teams, watched, path):
        stats = self._compute_tour_stats(use_teams, watched)
        half    = (len(stats) + 1) // 2
        left    = stats[:half]
        right   = stats[half:]

        while len(right) < len(left): right.append(["", "", None])
        split_stats = []
        for l, r in zip(left, right): split_stats.append([l[0], l[1], r[0], r[1]])

        df_tour = pd.DataFrame(split_stats, columns = ["Metric", "Value", "Metric", "Value"])
        self._export_png(df_tour, path, "Tour.png", "Tour Statistics")

    def _download_challonge_page(self, url: str) -> str:
        headers = {
            "accept"            : "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "accept-language"   : "en-US,en;q=0.9",
            "cache-control"     : "no-cache",
            "referer"           : "https://challonge.com/",
        }

        parsed = urlparse(url.strip())
        if not parsed.scheme: parsed = urlparse("https://" + url.strip())

        base        = urlunparse((parsed.scheme or "https", parsed.netloc, parsed.path.rstrip("/"), "", "", ""))
        variants    = [base, base + "/module?multiplier=1&match_width_multiplier=1&show_final_results=1"]

        for candidate_url in variants:
            for imp in ["chrome124", "chrome123", "chrome120"]:
                try:
                    res = requests.get(candidate_url, headers=headers, impersonate=imp, timeout=15)
                    if res.status_code == 200: return res.text

                except: continue

        raise RuntimeError("[!] Failed to fetch Challonge data: Blocked by Challonge")

    def _parse_challonge_display_leader(self, display_name: str):
        player_text = (display_name or "").split("|", 1)[0]
        pattern     = r"([^\s\[(|]+)(?:\s*\[(.*?)\])?(?:\s*\((-?\d+(?:\.\d+)?)\))?"
        ignored     = {"total", "guesses", "average", "avg", "="}

        for name, _, _ in re.findall(pattern, player_text):
            if not name or name.casefold() in ignored or re.fullmatch(r"-?\d+(?:\.\d+)?", name): continue
            return name

        return "Unknown Team"

    def _compute_team_rows(self, assigns, t1_lookup):
        if not getattr(self, 'use_teams', False): return pd.DataFrame()

        if not t1_lookup:
            for tid, players in self.rosters.items():
                if players:
                    sorted_players = sorted(list(players), key = str.lower)
                    t1_lookup[tid] = sorted_players[0]

        chal_matches = []

        if getattr(self, 'challonge_choice', 'No') == "Yes":
            codes_file  = self.tour_dir / FILE_CODES
            chal_link   = None

            if codes_file.exists():
                with open(codes_file, "r", encoding = "utf-8") as f:
                    for line in f:
                        if line.strip().startswith("http"):
                            chal_link = line.strip()
                            break

            if chal_link:
                try:
                    html = self._download_challonge_page(chal_link)
                    soup = BeautifulSoup(html, "lxml")

                    for script in soup.find_all("script"):
                        if script.string and "window._initialStoreState" in script.string:
                            match = re.search(r"window\._initialStoreState\['TournamentStore'\]\s*=\s*({.*?});", script.string, re.DOTALL)

                            if match:
                                data_str    = match.group(1).replace("'", '"')
                                data_str    = re.sub(r",\s*}", "}", data_str)
                                data_str    = re.sub(r",\s*]", "]", data_str)
                                s_data      = json.loads(data_str)

                                for _, rounds in s_data.get("matches_by_round", {}).items(): chal_matches.extend(rounds)
                                break

                except Exception as e: print(f"[!] Failed to fetch Challonge data: {e}")

        alias_map   = {}
        alias_path  = self.tour_dir / FILE_ALIAS

        if alias_path.exists():
            with open(alias_path, "r", encoding = "utf-8") as f:
                for line in f:
                    if "," in line:
                        parts = [p.strip().lower() for p in line.split(",")]
                        if len(parts) >= 2:
                            alias_map[parts[0]] = parts[1]
                            alias_map[parts[1]] = parts[0]

        res = []

        for tid in self.t_c_ps:
            leader_name     = t1_lookup.get(tid, f"Team {tid}")
            leader_variants = {leader_name.lower()}

            if leader_name.lower() in alias_map: leader_variants.add(alias_map[leader_name.lower()])

            wins_history                = defaultdict(list)
            losses_history              = defaultdict(list)
            ties_history                = defaultdict(list)
            w_count, l_count, t_count   = 0, 0, 0

            for m in chal_matches:
                scores = m.get("scores")
                if not isinstance(scores, list) or len(scores) < 2: continue

                try     : s1, s2 = int(scores[0]), int(scores[1])
                except  : continue

                p1_leader = self._parse_challonge_display_leader(m["player1"].get("display_name", ""))
                p2_leader = self._parse_challonge_display_leader(m["player2"].get("display_name", ""))

                p1_match = p1_leader.lower() in leader_variants or alias_map.get(p1_leader.lower()) in leader_variants
                p2_match = p2_leader.lower() in leader_variants or alias_map.get(p2_leader.lower()) in leader_variants

                if p1_match or p2_match:
                    if p1_match : my_score, opp_score, opp_name = s1, s2, p2_leader
                    else        : my_score, opp_score, opp_name = s2, s1, p1_leader

                    score_str = f"{my_score}-{opp_score}"
                    if my_score > opp_score:
                        w_count += 1
                        wins_history[opp_name].append(score_str)

                    elif my_score < opp_score:
                        l_count += 1
                        losses_history[opp_name].append(score_str)

                    else:
                        t_count += 1
                        ties_history[opp_name].append(score_str)

            h_payload = {
                "summary"       : f"{w_count}-{l_count}" if t_count == 0 else f"{w_count}-{l_count}-{t_count}",
                "wins"          : dict(wins_history),
                "losses"        : dict(losses_history),
                "ties"          : dict(ties_history),
                "total_matches" : w_count + l_count + t_count
            }

            t_overs = []

            for original_name in self.s_part:
                if original_name.lower() in assigns:
                    if assigns[original_name.lower()][0] == tid and self.c_counts[original_name] > 0: t_overs.append(self.p_overs_sum[original_name] / self.c_counts[original_name])
         
            t_elos = []

            for p in self.rosters[tid]:
                v = self.elo_map.get(p.lower())

                if v is not None:
                    try     : t_elos.append(float(v))
                    except  : pass

            row = {
                "Team Leader"   : leader_name,
                "Mean Elo"      : np.mean(t_elos),
                "Mean GR"       : np.mean(self.t_c_ps       [tid]) * 100,
                "Total 1/8s"    : self.t_solos              [tid],
                "Mean Over-8"   : np.mean(t_overs),
                "Rig Synergy"   : np.mean(self.t_on_syn     [tid]) * 100,
                "Off Synergy"   : np.mean(self.t_off_syn    [tid]) * 100,
                "Shared Rigs"   : np.mean(self.t_sh_rig     [tid]) * 100,
                "_tid"          : tid,
                "_history"      : h_payload
            }

            if getattr(self, 'text_var_wlt', 'No') == 'Yes' or h_payload["total_matches"] > 0:
                row["Win Record"]   = f"{w_count}-{l_count}" if t_count == 0 else h_payload["summary"]
                tot                 = w_count + l_count + t_count
                row["_win_pct"]     = (w_count / tot) if tot > 0 else -1.0

            res.append(row)

        df = pd.DataFrame(res).sort_values(by = ["Mean GR", "Mean Elo"], ascending = [False, True])
        return df

    def _create_team_png(self, assigns, t1_lookup, path):
        self.text_var_wlt = "Yes" 
        df = self._compute_team_rows(assigns, t1_lookup)

        if "_win_pct" in df.columns : df_png = df.drop(columns = ["_tid", "_history", "_win_pct"])
        else                        : df_png = df.drop(columns = ["_tid", "_history"])

        if "Win Record" in df_png.columns:
            if (df_png["Win Record"].astype(str) == "0-0-0").all() or (df_png["Win Record"].astype(str) == "0-0").all(): df_png = df_png.drop(columns = ["Win Record"])

        watched_valid = self.missing_list_count <= THRESH_WTCH

        if not watched_valid:
            df_png = df_png.drop(columns = ["Rig Synergy", "Off Synergy", "Shared Rigs"], errors = "ignore")
            num_cols = ["Mean Elo", "Mean GR", "Mean Over-8"]

        else: num_cols = ["Mean Elo", "Mean GR", "Mean Over-8", "Rig Synergy", "Off Synergy", "Shared Rigs"]

        for c in num_cols: df_png[c] = pd.to_numeric(df_png[c], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

        self._export_png(df_png, path, "Team.png", "Team Statistics")
        self.text_var_wlt = "No"

    def _compute_tier_rows(self, assigns, has_chanting_songs):
        rows1, rows2 = [], []

        for tr in ["1", "2", "3", "4"]:
            tp = [n for n in self.s_part if n.lower() in assigns and assigns[n.lower()][1] == tr]
            if not tp: continue

            row1 = {"Tier": tr}
            row2 = {"Tier": tr}

            gen_players, atk_players, blk_players, con_players, spd_players, chn_players = [], [], [], [], [], []

            for p in tp:
                cor = self.c_counts[p]
                tot = self.s_part[p]
                tim = self.p_answer_times.get(p, [])
                chc = self.p_chan_c[p]
                cht = self.p_chan_s[p]

                gen = 100 * cor / tot if tot else 0.0
                atk = self.p_pts[p]
                blk = self.p_blks[p]
                con = 100 * (atk + blk) / cor if cor else 0.0
                spd = np.median(tim) if tim else np.nan
                chn = 100 * chc / cht if cht else 0.0

                gen_players.append({"player": p, "value": gen})
                atk_players.append({"player": p, "value": atk})
                blk_players.append({"player": p, "value": blk})
                con_players.append({"player": p, "value": con})

                if pd.notnull(spd)      : spd_players.append({"player": p, "value": spd})
                if has_chanting_songs   : chn_players.append({"player": p, "value": chn})

            gen_players.sort(key = lambda x: x["value"], reverse = True)
            atk_players.sort(key = lambda x: x["value"], reverse = True)
            blk_players.sort(key = lambda x: x["value"], reverse = True)
            con_players.sort(key = lambda x: x["value"], reverse = True)
            spd_players.sort(key = lambda x: x["value"], reverse = False)

            if chn_players: chn_players.sort(key = lambda x: x["value"], reverse = True)

            row1["GR"]                  = f"{gen_players[0]['player']} ({gen_players[0]['value']:.2f})" if gen_players                                  else "N/A"
            row1["Lives Taken"]         = f"{atk_players[0]['player']} ({atk_players[0]['value']:g})"   if atk_players                                  else "N/A"
            row1["Lives Saved"]         = f"{blk_players[0]['player']} ({blk_players[0]['value']:g})"   if blk_players                                  else "N/A"
            row2["Contribution Rate"]   = f"{con_players[0]['player']} ({con_players[0]['value']:.2f})" if con_players                                  else "N/A"
            row2["Median Time"]         = f"{spd_players[0]['player']} ({spd_players[0]['value']:.2f})" if spd_players                                  else "N/A"
            row2["Chant GR"]            = f"{chn_players[0]['player']} ({chn_players[0]['value']:.2f})" if chn_players and chn_players[0]['value'] > 0  else ""

            row1["gen_val"] = gen_players[0]['value'] if gen_players else 0.0
            row1["atk_val"] = atk_players[0]['value'] if atk_players else 0.0
            row1["blk_val"] = blk_players[0]['value'] if blk_players else 0.0
            row2["con_val"] = con_players[0]['value'] if con_players else 0.0
            row2["spd_val"] = spd_players[0]['value'] if spd_players else 0.0
            row2["chn_val"] = chn_players[0]['value'] if chn_players else 0.0

            row1["_players"] = {"gen": gen_players, "atk": atk_players, "blk": blk_players}
            row2["_players"] = {"con": con_players, "spd": spd_players, "chn": chn_players}

            rows1.append(row1)
            rows2.append(row2)

        return rows1, rows2

    def _create_tier_png(self, assigns, path, has_chanting_songs):
        rows1, rows2    = self._compute_tier_rows(assigns, has_chanting_songs)
        valid_spd_rows  = [i for i, r in enumerate(rows2) if pd.notnull(r["spd_val"])]
        valid_chn_rows  = [i for i, r in enumerate(rows2) if r["chn_val"] is not None]

        best_gen_idx = len(rows1) - 1 - max(range(len(rows1)),  key = lambda i: rows1[::-1][i]["gen_val"]) if rows1             else None
        best_atk_idx = len(rows1) - 1 - max(range(len(rows1)),  key = lambda i: rows1[::-1][i]["atk_val"]) if rows1             else None
        best_blk_idx = len(rows1) - 1 - max(range(len(rows1)),  key = lambda i: rows1[::-1][i]["blk_val"]) if rows1             else None
        best_con_idx = len(rows2) - 1 - max(range(len(rows2)),  key = lambda i: rows2[::-1][i]["con_val"]) if rows2             else None
        best_spd_idx = len(rows2) - 1 - min(valid_spd_rows,     key = lambda i: rows2[::-1][i]["spd_val"]) if valid_spd_rows    else None
        best_chn_idx = len(rows2) - 1 - max(valid_chn_rows,     key = lambda i: rows2[::-1][i]["chn_val"]) if valid_chn_rows    else None

        html_parts = []
        html_parts.append("<tr><th>Tier</th><th>Guess Rate</th><th>Lives Taken</th><th>Lives Saved</th></tr>")

        style_hl = f" style='background-color: {COLOR_2}; color: white; font-weight: bold;'"

        for idx, row in enumerate(rows1):
            s_gen = style_hl if idx == best_gen_idx else ""
            s_atk = style_hl if idx == best_atk_idx else ""
            s_blk = style_hl if idx == best_blk_idx else ""

            html_parts.append(
                f"<tr>"
                    f"<td><b>{row['Tier']}</b></td>"
                    f"<td{s_gen}>{row['GR']}</td>"
                    f"<td{s_atk}>{row['Lives Taken']}</td>"
                    f"<td{s_blk}>{row['Lives Saved']}</td>"
                f"</tr>"
            )

        html_parts.append(
            "<tr>"
                "<th style='border-top: 3px solid black;'>Tier</th>"
                "<th style='border-top: 3px solid black;'>Contribution Rate</th>"
                "<th style='border-top: 3px solid black;'>Median Time</th>"
                "<th style='border-top: 3px solid black;'>Chanting GR</th>"
            "</tr>"
        )

        for idx, row in enumerate(rows2):
            s_con = style_hl if idx == best_con_idx else ""
            s_spd = style_hl if idx == best_spd_idx else ""
            s_chn = style_hl if idx == best_chn_idx else ""

            html_parts.append(
                f"<tr>"
                    f"<td><b>{row['Tier']}</b></td>"
                    f"<td{s_con}>{row['Contribution Rate']}</td>"
                    f"<td{s_spd}>{row['Median Time']}</td>"
                    f"<td{s_chn}>{row['Chant GR']}</td>"
                f"</tr>"
            )

        html_table_content = "".join(html_parts)

        full = f"""<html>
            <head>
                <style>
                    body {{
                        font-family         : 'Segoe UI';
                        background          : white;
                        display             : inline-block;
                        margin              : 0
                    }}

                    h2 {{
                        margin              : 0 0 10px 0;
                        font-size           : 40px;
                        text-align          : center
                    }}

                    table {{
                        border-collapse     : collapse;
                        width               : auto;
                        border              : 3px solid black
                    }}

                    th {{
                        font-weight         : bold;
                        font-size           : 25px;
                        text-align          : center;
                        padding             : 10px;
                        border              : 1px solid black;
                        border-bottom       : 3px solid black;
                        background-color    : #f0f0f0
                    }}

                    td {{
                        font-size           : 25px;
                        text-align          : center;
                        padding             : 10px;
                        border              : 1px solid black
                    }}

                    tr:nth-child(even) {{background-color: #f0f0f0}}
                </style>
            </head>
            <body>
                <h2>Tier Statistics</h2>
                <table>{html_table_content}</table>
            </body>
        </html>"""

        if not self.browser_path: return
        hti = Html2Image(size = (2000, 2000), browser_executable = self.browser_path, output_path = str(path), custom_flags = ['--log-level=3', '--silent'])
        hti.screenshot(html_str = full, save_as = "Tier.png")

        try     : trim_whitespace(path / "Tier.png")
        except  : pass

    def _create_scatter_png(self, path, list_mode = False, elo_map = None):
        configs = []

        if list_mode:
            plist_l = [n for n in self.s_part if self.p_l_corr[n]]

            if plist_l:
                x_vals_l    = [np.mean(self.p_l_corr[name]) for name in plist_l]
                y_vals_l    = [np.median(self.p_l_vint[name]) if self.p_l_vint[name] else np.nan for name in plist_l]
                valid_l     = [(p, x, y) for p, x, y in zip(plist_l, x_vals_l, y_vals_l) if not np.isnan(y)]
                
                if valid_l:
                    plist_l, x_vals_l, y_vals_l = zip(*valid_l)
                    plist_l, x_vals_l, y_vals_l = list(plist_l), list(x_vals_l), list(y_vals_l)
                    
                    rig_rates   = [self.p_rigs      [name] / self.s_part[name] if self.s_part[name] else 0 for name in plist_l]
                    grid_grs    = [self.p_rigs_h    [name] / self.p_rigs[name] if self.p_rigs[name] else 0 for name in plist_l]

                    scale_l = 1.00 if len(plist_l) <= THRESH_PLYR else (0.75 if len(plist_l) <= THRESH_PLYR + 8 else 0.50)
                    sizes_l = [(rate * scale_l) ** 2 * 10000 for rate in rig_rates]

                    cmap_l = mc.LinearSegmentedColormap.from_list("rig_gr_cmap", [
                        (0.0, COLOR_0),
                        (0.7, COLOR_0),
                        (0.8, COLOR_1),
                        (0.9, COLOR_2),
                        (1.0, COLOR_2)
                    ])
                    
                    configs.append({
                        "filename"          : "List.png",
                        "title"             : "List Statistics",
                        "plist"             : plist_l,
                        "x_vals"            : x_vals_l,
                        "y_vals"            : y_vals_l,
                        "sizes"             : sizes_l,
                        "colors"            : grid_grs,
                        "cmap"              : cmap_l,
                        "vmin"              : 0.0,
                        "vmax"              : 1.0,
                        "cbar_label"        : "Rig GR",
                        "cbar_ticks"        : [0, 0.7, 0.8, 0.9, 1],
                        "cbar_ticklabels"   : ['0', '70', '80', '90', '100'],
                        "labelpad"          : -35
                    })

        plist_g = [n for n in self.s_part if self.c_counts[n] > 0]

        if plist_g:
            x_vals_g    = [self.p_overs_sum[name] / self.c_counts[name] for name in plist_g]
            y_vals_g    = [np.median(self.p_c_vint[name]) if self.p_c_vint[name] else np.nan for name in plist_g]
            valid_g     = [(p, x, y) for p, x, y in zip(plist_g, x_vals_g, y_vals_g) if not np.isnan(y)]
            
            if valid_g:
                plist_g, x_vals_g, y_vals_g = zip(*valid_g)
                plist_g, x_vals_g, y_vals_g = list(plist_g), list(x_vals_g), list(y_vals_g)
                
                gr_vals = [self.c_counts[name] / self.s_part[name] if self.s_part[name] else 0 for name in plist_g]
                if elo_map is None: elo_map = {}
                uf_pool, el_pool = [], []

                valid_elos  = [float(v) for v in elo_map.values() if str(v).replace('.', '', 1).isdigit() or (str(v).startswith('-') and str(v)[1:].replace('.', '', 1).isdigit())]
                avg_rank    = np.mean(valid_elos) if valid_elos else 1.0

                for name in plist_g:
                    tot         = self.s_part[name]
                    uf_scaled   = (self.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0
                    elo_val     = elo_map.get(name.lower(), 0.0)

                    try     : elo = float(elo_val)
                    except  : elo = 0.0

                    uf_pool.append(uf_scaled)
                    el_pool.append(elo)

                if uf_pool and el_pool:
                    els, ufs = np.array(el_pool), np.array(uf_pool)

                    if len(els) > 1 and np.var(els) > 0:
                        slope, intercept = np.polyfit(els, ufs, 1)
                        expected_ufs = slope * els + intercept

                    else: expected_ufs = np.array([np.mean(ufs)] * len(ufs))

                    residuals   = ufs - expected_ufs
                    res_std     = np.std(residuals) if np.std(residuals) > 0 else 1
                    norm_perf   = [1 / (1 + np.exp(SCALE_PERF * (res / res_std))) for res in residuals]

                else: norm_perf = [0.5] * len(plist_g)

                scale_g = 1.00 if len(plist_g) <= THRESH_PLYR else (0.75 if len(plist_g) <= THRESH_PLYR + 8 else 0.50)
                sizes_g = [(rate * scale_g) ** 2 * 10000 for rate in gr_vals]
                cmap_g  = mc.LinearSegmentedColormap.from_list("guess_uf_elo_cmap", [(0, COLOR_0), (0.5, COLOR_1), (1, COLOR_2)])
                
                configs.append({
                    "filename"          : "Guess.png",
                    "title"             : "Guess Statistics",
                    "plist"             : plist_g,
                    "x_vals"            : x_vals_g,
                    "y_vals"            : y_vals_g,
                    "sizes"             : sizes_g,
                    "colors"            : norm_perf,
                    "cmap"              : cmap_g,
                    "vmin"              : 0.0,
                    "vmax"              : 1.0,
                    "cbar_label"        : "Score",
                    "cbar_ticks"        : [0, 1],
                    "cbar_ticklabels"   : ['0', '100'],
                    "labelpad"          : -37.5
                })

        if not configs: return
        all_x, all_y = [], []

        for cfg in configs:
            all_x.extend(cfg["x_vals"])
            all_y.extend(cfg["y_vals"])

        if not all_x or not all_y: return

        x_min = math.floor  ((min   (all_x) - 0.5) * 2) / 2
        x_max = math.ceil   ((max   (all_x) + 0.5) * 2) / 2
        y_min = math.floor  (min    (all_y) - 1.0)
        y_max = math.ceil   (max    (all_y) + 1.0)

        while True:
            r = y_max - y_min

            if r % 4 == 0:
                y_stp = r // 4
                break

            elif r % 3 == 0:
                y_stp = r // 3
                break

            if r % 2 != 0 or y_max >= 2026  : y_min -= 1
            else                            : y_max += 1

        x_center = (x_min + x_max) / 2
        y_center = (y_min + y_max) / 2

        for cfg in configs:
            fig, ax = plt.subplots(figsize = (10, 10))
            sc      = ax.scatter(
                cfg["x_vals"], cfg["y_vals"],
                s           = cfg["sizes"],
                c           = cfg["colors"], 
                cmap        = cfg["cmap"],
                vmin        = cfg["vmin"],
                vmax        = cfg["vmax"], 
                edgecolors  = 'black',
                alpha       = 0.95,
                zorder      = 3
            )

            points          = np.column_stack((cfg["x_vals"], cfg["y_vals"]))
            x_range         = max(cfg["x_vals"]) - min(cfg["x_vals"]) if max(cfg["x_vals"]) != min(cfg["x_vals"]) else 1
            y_range         = max(cfg["y_vals"]) - min(cfg["y_vals"]) if max(cfg["y_vals"]) != min(cfg["y_vals"]) else 1
            norm_points     = np.column_stack((points[:, 0] / x_range, points[:, 1] / y_range))
            center_of_mass  = np.median(norm_points, axis = 0)
            distances       = np.linalg.norm(norm_points - center_of_mass, axis = 1)
            pack_mask       = distances < np.percentile(distances, 75)
            pack_points     = points[pack_mask]

            if len(pack_points) >= 3:
                try:
                    hull        = ConvexHull(pack_points)
                    hull_points = pack_points[hull.vertices]
                    hull_points = np.vstack([hull_points, hull_points[0]])

                    ax.plot(hull_points[:, 0], hull_points[:, 1], color = 'black', zorder = 1, linewidth = 0.5, linestyle = '-')

                except Exception: pass

            ax.set_xlim(x_min, x_max)
            ax.set_ylim(y_min, y_max)

            x_mnc, x_mxf, x_stp = math.ceil(x_min), math.floor(x_max), 1

            ax.set_xticks(range(x_mnc,      x_mxf + x_stp, x_stp))
            ax.set_yticks(range(y_min + 1,  y_max + y_stp, y_stp))

            texts = []

            for name, x, y in zip(cfg["plist"], cfg["x_vals"], cfg["y_vals"]):
                label = self._get_player_acronym(name)
                if not label: continue

                ha_align = "left"   if x >= x_center else "right"
                va_align = "bottom" if y >= y_center else "top"

                texts.append(ax.text(
                    x, y, label,
                    size        = 20,
                    weight      = "bold",
                    fontname    = "Segoe UI",
                    ha          = ha_align,
                    va          = va_align
                ))

            if texts: adjust_text(
                texts,
                ax                      = ax,
                objects                 = sc,
                avoid_self              = True,
                add_objects_to_edges    = True,
                force_text              = (1.00, 1.00),
                force_objects           = (1.00, 1.00),
                expand                  = (2.00, 2.00),
                arrowprops              = dict(arrowstyle = "-", color = 'black', shrinkA = 15)
            )

            ax.set_title    (cfg["title"],  weight = 'bold', fontname = "Segoe UI", fontsize = 50, pad      = 15)
            ax.set_xlabel   ("Over-8",      weight = 'bold', fontname = "Segoe UI", fontsize = 25, labelpad = 5)
            ax.set_ylabel   ("Vintage",     weight = 'bold', fontname = "Segoe UI", fontsize = 25, labelpad = 5)

            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, _: str(int(val))))
            plt.setp(ax.get_yticklabels(), rotation = 90, va = 'center')

            ax.tick_params(axis = 'x', which = 'both', length = 0, labelsize = 20, pad = 5)
            ax.tick_params(axis = 'y', which = 'both', length = 0, labelsize = 20, pad = 2.5)

            cbar = fig.colorbar(sc, ax = ax, pad = 0.005, aspect = 40, ticks = cfg["cbar_ticks"])
            cbar.set_label(cfg["cbar_label"], weight = 'bold', fontname = "Segoe UI", fontsize = 25, labelpad = cfg["labelpad"])

            cbar.ax.set_yticklabels(cfg["cbar_ticklabels"])
            cbar.ax.tick_params(labelsize = 20, length = 0)

            ax.text(0.01, 0.99, "New\nHard", transform = ax.transAxes, color = "black", fontsize = 15, va = "top",       ha = "left",    weight = "bold", alpha = 0.75)
            ax.text(0.99, 0.99, "New\nEasy", transform = ax.transAxes, color = "black", fontsize = 15, va = "top",       ha = "right",   weight = "bold", alpha = 0.75)
            ax.text(0.01, 0.01, "Old\nHard", transform = ax.transAxes, color = "black", fontsize = 15, va = "bottom",    ha = "left",    weight = "bold", alpha = 0.75)
            ax.text(0.99, 0.01, "Old\nEasy", transform = ax.transAxes, color = "black", fontsize = 15, va = "bottom",    ha = "right",   weight = "bold", alpha = 0.75)

            ax.grid(False)

            plt.tight_layout    ()
            plt.savefig         (path / cfg["filename"], dpi = 100)
            plt.close           (fig)

            try     : trim_whitespace(path / cfg["filename"])
            except  : pass

    def _create_song_png(self, path):
        diffs       = [s["difficulty"] for s in self.song_data]
        max_diff    = max(diffs) if diffs else 0

        if max_diff < 40:
            num_x   = 4
            num_y   = 4
            font_sz = 90

        else:
            num_x   = 5
            num_y   = 5
            font_sz = 75

        counts      = np.zeros((num_y, num_x), dtype = int)
        over8_sums  = np.zeros((num_y, num_x), dtype = float)

        for s in self.song_data:
            vint = int(s["vintage"])
            if vint == 0: continue

            diff    = s["difficulty"]
            x_idx   = min(int(math.floor(diff / 10)), num_x - 1)
            
            if num_y == 4:
                if vint < 2000  : y_idx = 0
                else            : y_idx = min(int(math.floor((vint - 2000) / 10)) + 1, 3)

            else: y_idx = min(max(int(math.floor((vint - 1980) / 10)), 0), 4)

            counts      [y_idx, x_idx] += 1
            over8_sums  [y_idx, x_idx] += s["correct_count"]

        fig, ax     = plt.subplots(figsize = (10, 10))
        cmap_song   = mc.LinearSegmentedColormap.from_list("song_cmap", [
            (0.000, COLOR_0),
            (0.375, COLOR_1),
            (0.625, COLOR_2),
            (1.000, COLOR_2)
        ])

        for y_idx in range(num_y):
            for x_idx in range(num_x):
                count = counts[y_idx, x_idx]

                if count == 0: facecolor = 'white'

                else:
                    avg_over8 = over8_sums[y_idx, x_idx] / count
                    facecolor = cmap_song(avg_over8 / 8.0)

                rect = plt.Rectangle((x_idx, y_idx), 1, 1, facecolor = facecolor, edgecolor = 'none')
                ax.add_patch(rect)

                if count > 0: ax.text(
                    x_idx + 0.5, y_idx + 0.45, str(count),
                    ha          = 'center',
                    va          = 'center',
                    color       = 'white',
                    weight      = 'bold',
                    fontsize    = font_sz,
                    fontname    = "Segoe UI"
                )

        ax.set_xlim(0, num_x)
        ax.set_ylim(0, num_y)

        ax.set_xticks(np.arange(1, num_x))
        ax.set_yticks(np.arange(1, num_y))

        if num_x == 4:
            p_labels = ["10",   "20",   "30"]
            y_labels = ["2000", "2010", "2020"]

        else:
            p_labels = ["10",   "20",   "30",   "40"]
            y_labels = ["1990", "2000", "2010", "2020"]

        ax.set_xticklabels(p_labels, fontname = "Segoe UI", fontsize = 20)
        ax.set_yticklabels(y_labels, fontname = "Segoe UI", fontsize = 20, rotation = 90, va = 'center')

        ax.set_title    ("Song Statistics", weight = 'bold', fontname = "Segoe UI", fontsize = 50, pad      = 15)
        ax.set_xlabel   ("Difficulty",      weight = 'bold', fontname = "Segoe UI", fontsize = 25, labelpad = 5)
        ax.set_ylabel   ("Vintage",         weight = 'bold', fontname = "Segoe UI", fontsize = 25, labelpad = 5)

        ax.tick_params(axis = 'x', which = 'both', length = 0, pad = 5)
        ax.tick_params(axis = 'y', which = 'both', length = 0, pad = 2.5)

        ax.grid(False)
        norm = mc.Normalize(vmin = 0.0, vmax = 8.0)

        sm = plt.cm.ScalarMappable(cmap = cmap_song, norm = norm)
        sm.set_array([])

        cbar = fig.colorbar(sm, ax = ax, pad = 0.005, aspect = 40, ticks = [0, 3, 5, 8])
        cbar.set_label("Over-8", weight = 'bold', fontname = "Segoe UI", fontsize = 25, labelpad = -12.5)

        cbar.ax.set_yticklabels(['0', '3', '5', '8'])
        cbar.ax.tick_params(labelsize = 20, length = 0)

        plt.tight_layout    ()
        plt.savefig         (path / "Song.png", dpi = 100)
        plt.close           (fig)

        try     : trim_whitespace(path / "Song.png")
        except  : pass

    def _get_dashboard_data(self):
        diffs       = [s["difficulty"] for s in self.song_data]
        max_diff    = max(diffs) if diffs else 0
        num_x       = 8 if max_diff < 40 else 9
        num_y       = 8 if max_diff < 40 else 9

        return (
            self.player_song_details, 
            self.tour_song_details, 
            self.team_song_details, 
            self.raw_vintage_by_guess, 
            self.raw_vintage_by_list, 
            self.matrix_song_details, 
            num_x, 
            num_y
        )

    def _render_dashboard_player(self, sorted_players, active, t_labels, watched, player_song_details, df_base):
        rows, eligibility, borders = [], [], []

        for _, name in enumerate(sorted_players):
            row_data    = df_base.loc[df_base["Player"].str.startswith(name)].iloc[0] if any(df_base["Player"].str.startswith(name)) else None
            target     = self.exp_map.get(name, self.base_exp)
            d_name     = name
            sub_hover  = ""

            if name in self.new_players  : d_name += " ★"

            if target != "ignore" and target < self.base_exp:
                if name.lower() in self.main_roster_names:
                    d_name  +=  " ▼"
                    subs    =   self.sub_relations.get(name.casefold(), [])

                    if subs: sub_hover = f"Subbed by {', '.join(subs)}"

                else:
                    d_name  +=  " ▲"
                    orig    =   self.sub_relations.get(name.lower(), [])

                    if orig: sub_hover = f"Subbing for {orig[0]}"

            is_eligible = not ("▼" in d_name or "▲" in d_name)
            eligibility.append(is_eligible)

            row = {"Player": {"count": d_name, "details": [sub_hover] if sub_hover else []}}

            if self.use_teams:
                t_info      = self.assignments  .get(name.lower(),  (None, "N/A"))
                team_leader = self.t1_lookup    .get(t_info[0],     "N/A") if t_info[0] is not None else "N/A"

                row["Team"] = team_leader
                row["Tier"] = t_info[1]

                try     : row["Elo"] = float(self.elo_map.get(name.lower(), np.nan))
                except  : row["Elo"] = np.nan

            if row_data is not None:
                tot, cor = self.s_part[name], self.c_counts[name]
                for key_details in ["Overall", "Type 1", "Type 2", "Type 3", "Chant"]: self.player_song_details[name][key_details].sort(key = lambda s: s[2:].strip().lower())

                row.update({
                    "GR"            : {"count": float(row_data["GR"] * 100), "details": [f"{cor}/{tot}"] + self.player_song_details[name]["Overall"]},
                    "GR Δ"          : float (row_data['GR Δ'])          if pd.notnull(row_data.get("GR Δ"))     else np.nan,
                    "UF"            : float (row_data["UF"])            if "UF"     in row_data                 else np.nan,
                    "UF Δ"          : float (row_data['UF Δ'])          if pd.notnull(row_data.get("UF Δ"))     else np.nan,
                    "Score"         : float (row_data["Score"])         if "Score"  in row_data                 else np.nan,
                    "1/8s"          : int   (row_data["1/8s"]),
                    "2/8s"          : int   (row_data["2/8s"]),
                    "7/8s"          : int   (row_data["7/8s"]),
                    "Mean Over-8"   : float (row_data["Mean Over-8"])   if pd.notnull(row_data["Mean Over-8"])  else np.nan
                })

                if self.use_teams: row.update({"Lives Taken": int(row_data["Lives Taken"]), "Lives Saved": int(row_data["Lives Saved"])})
                
                for tid in active:
                    seen        = self.p_type_s[name][tid]
                    succ        = self.p_type_c[name][tid]
                    t_key       = t_labels[tid].split(" ")[0] 
                    delta_key   = f"{t_key} Δ"

                    row[t_labels[tid]] = {
                        "count"     : float(row_data[t_labels[tid]] * 100),
                        "details"   : [f"{succ}/{seen}"] + self.player_song_details[name][f"Type {tid}"]
                    } if pd.notnull(row_data[t_labels[tid]]) else np.nan

                    row[delta_key] = float(row_data[delta_key]) if pd.notnull(row_data.get(delta_key)) else np.nan

                if watched:
                    succ_rig    = self.p_rigs_h [name]
                    tot_rig     = self.p_rigs   [name]
                    succ_off    = self.c_counts [name] - succ_rig
                    tot_off     = self.s_part   [name] - tot_rig

                    self.player_song_details[name]["Rigs"].sort(key = lambda s: s[2:].strip().lower())

                    rig_song_lines  = {s[2:] for s in self.player_song_details[name]["Rigs"]}
                    off_details     = [s for s in self.player_song_details[name]["Overall"] if s[2:] not in rig_song_lines]

                    row.update({
                        "Rigs"          : int   (row_data["Rigs"]),
                        "Rig Rate"      : float (row_data["Rig Rate"]                   * 100),
                        "Solo Rigs"     : int   (row_data["Solo Rigs"]),
                        "Solo Rig Rate" : float (row_data["Solo Rig Rate"]              * 100),
                        "Rig Over-8"    : float (row_data["Rig Over-8"])                                                                                                        if pd.notnull(row_data["Rig Over-8"])   else np.nan,
                        "Over-8 Δ"      : float (row_data["Over-8 Δ"])                                                                                                          if pd.notnull(row_data["Over-8 Δ"])     else np.nan,
                        "Rig GR"        : {"count": float(row_data["Rig GR"]            * 100), "details": [f"{succ_rig}/{tot_rig}"] + self.player_song_details[name]["Rigs"]}  if pd.notnull(row_data["Rig GR"])       else np.nan,
                        "Off GR"        : {"count": float(row_data["Off GR"]            * 100), "details": [f"{succ_off}/{tot_off}"] + off_details}                             if pd.notnull(row_data["Off GR"])       else np.nan,
                        "Rig Δ"         : float (row_data["Rig Δ"]                      * 100),
                    })

                h_diffs = self.p_hit_diff.get(name, [])
                h_vints = self.p_hit_vint.get(name, [])

                if h_diffs:
                    row["Mean Difficulty Hit"] = {
                        "count"     : float(np.mean(h_diffs)),
                        "details"   : [
                            f"Minimum: {float(np.min(h_diffs)):.2f}",
                            f"Median: {float(np.median(h_diffs)):.2f}",
                            f"Maximum: {float(np.max(h_diffs)):.2f}",
                            f"Standard Deviation: {float(np.std(h_diffs)):.2f}"
                        ]
                    }

                else: row["Mean Difficulty Hit"] = np.nan

                if h_vints:
                    row["Median Vintage Hit"] = {
                        "count"     : float(np.median(h_vints)),
                        "details"   : [
                            f"Minimum: {float(np.min(h_vints))}",
                            f"Mean: {float(np.mean(h_vints)):.2f}",
                            f"Maximum: {float(np.max(h_vints))}",
                            f"Standard Deviation: {float(np.std(h_vints)):.2f}"
                        ]
                    }

                else: row["Median Vintage Hit"] = np.nan

                times = self.p_answer_times.get(name, [])

                if times:
                    t_min   = float(np.min(times))
                    t_mean  = float(np.mean(times))
                    t_max   = float(np.max(times))
                    t_std   = float(np.std(times))
                    t_det   = [
                        f"Minimum: {t_min:.2f}",
                        f"Mean: {t_mean:.2f}",
                        f"Maximum: {t_max:.2f}",
                        f"Standard Deviation: {t_std:.2f}"
                    ]

                    row["Median Time"] = {"count": float(row_data["Median Time"]), "details": t_det}

                else: row["Median Time"] = np.nan

                seen_chan = self.p_chan_s[name]
                succ_chan = self.p_chan_c[name]

                row["Chant GR"] = {
                    "count"     : float(row_data["Chant GR"] * 100),
                    "details"   : [f"{succ_chan}/{seen_chan}"] + self.player_song_details[name]["Chant"]
                } if pd.notnull(row_data["Chant GR"]) else np.nan

            for key in ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Solo Rigs"]:
                if key not in row               : continue
                if key in ["Rigs", "Solo Rigs"] : player_song_details[name][key].sort(key = lambda s: s[2:].strip().lower())
                else                            : player_song_details[name][key].sort(key = str.lower)

                row[key] = {"count": row[key], "details": player_song_details[name][key]}

            rows.append(row)

        df_players          = pd.DataFrame(rows)
        delta_json_fields   = ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]

        for c_field in delta_json_fields:
            if c_field in df_players.columns:
                if df_players[c_field].isna().all() or (df_players[c_field].astype(str) == "nan").all(): df_players = df_players.drop(columns = [c_field])

        if "GR" in df_players.columns and "Eru" not in self.tour_label:
            if self.val_str == "default":
                if      self.tour_label == "Watched 2+8s"               : th_val = "25, 20, 15, 10, 5"
                elif    self.tour_label in ["Watched", "QuagWatched"]   : th_val = "28, 18, 12, 6"
                elif    self.tour_label in ["Usual", "Quagsual"]        : th_val = "28, 19, 8"
                elif    watched                                         : th_val = "28, 18, 12, 6"
                else                                                    : th_val = "28, 19, 8"

            else: th_val = self.val_str

            try     : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
            except  : th = [28.0, 18.0, 12.0, 6.0]

            gv = df_players["GR"].map(lambda x: x["count"] if isinstance(x, dict) else x).tolist()

            for t in th:
                f_idx = -1

                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i

                if f_idx != -1 and f_idx < len(df_players) - 1: borders.append(int(f_idx))

        desc_cols = [
            "Elo",
            "GR",
            "GR Δ",
            "UF",
            "UF Δ",
            "Score",
            "1/8s",
            "2/8s",
            "Lives Taken",
            "Lives Saved",
            "OP GR",
            "OP Δ",
            "ED GR",
            "ED Δ",
            "IN GR",
            "IN Δ",
            "Rigs",
            "Rig Rate",
            "Solo Rigs",
            "Solo Rig Rate",
            "Over-8 Δ",
            "Rig GR",
            "Off GR",
            "Rig Δ",
            "Median Vintage Hit",
            "Chant GR",
        ]

        asc_cols    = ["7/8s", "Median Time", "Mean Over-8", "Rig Over-8", "Mean Difficulty Hit"]
        int_cols    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs", "Solo Rigs"]
        rate_cols   = ["GR", "OP GR", "ED GR", "IN GR", "Chant GR", "Rig GR", "Off GR"]
        stats_hl    = {}

        elo_ser     = df_players["Elo"]     .fillna(0.0) if "Elo" in df_players.columns else pd.Series(0.0, index = df_players.index)
        gr_ser      = df_players["GR"]      .map(lambda x: x["count"] if isinstance(x, dict) else x).fillna(0.0)
        rig_ser     = df_players["Rigs"]    .map(lambda x: x["count"] if isinstance(x, dict) else x).fillna(0.0) if "Rigs" in df_players.columns else pd.Series(0.0, index = df_players.index)

        mask_series = pd.Series(eligibility, index = df_players.index)

        for col in df_players.columns:
            string_cols = ["Team", "Tier"]
            dict_cols   = ["Mean Difficulty Hit", "Median Vintage Hit", "Median Time"]

            if col in string_cols: continue

            if col in desc_cols or col in asc_cols:
                if col in int_cols or col in rate_cols or col in dict_cols  : num = df_players[col].map(lambda x: x["count"] if isinstance(x, dict) else x)
                else                                                        : num = df_players[col]

                num     = pd.to_numeric(num, errors = 'coerce')
                el_num  = num[mask_series].dropna() if col in int_cols else num.dropna()

                if not num.dropna().empty:
                    if col in desc_cols:
                        best_val    = num.dropna().max()
                        worst_val   = el_num.min() if not el_num.empty else None

                    else:
                        best_val = num.dropna().min()

                        if col == "Median Time":
                            under_limit = el_num[el_num < THRESH_TIME]
                            worst_val   = under_limit.max() if not under_limit.empty else None

                        else: worst_val = el_num.max() if not el_num.empty else None

                    best_b_idx  = num       [num    == best_val]    .index if pd.notnull(best_val)  else pd.Index([])
                    worst_b_idx = el_num    [el_num == worst_val]   .index if pd.notnull(worst_val) else pd.Index([])

                    if col == "Solo Rigs":
                        best_idx    = int(rig_ser.loc[best_b_idx]   .idxmin()) if not best_b_idx    .empty else None
                        worst_idx   = int(rig_ser.loc[worst_b_idx]  .idxmax()) if not worst_b_idx   .empty else None

                    elif col == "Solo Rig Rate":
                        best_idx    = int(rig_ser.loc[best_b_idx]   .idxmax()) if not best_b_idx    .empty else None
                        worst_idx   = int(rig_ser.loc[worst_b_idx]  .idxmax()) if not worst_b_idx   .empty else None

                    elif col in ["Elo"]:
                        best_idx    = int(best_b_idx    [0]) if not best_b_idx  .empty else None
                        worst_idx   = int(worst_b_idx   [0]) if not worst_b_idx .empty else None

                    elif col in ["OP GR", "ED GR", "IN GR", "Chant GR"]:
                        best_idx    = int(gr_ser.loc[best_b_idx]    .idxmin()) if not best_b_idx    .empty else None
                        worst_idx   = int(gr_ser.loc[worst_b_idx]   .idxmax()) if not worst_b_idx   .empty else None

                    elif col == "Rig GR":
                        best_idx    = int(rig_ser.loc[best_b_idx]   .idxmax()) if not best_b_idx    .empty else None
                        worst_idx   = int(elo_ser.loc[worst_b_idx]  .idxmax()) if not worst_b_idx   .empty else None

                    else:
                        best_idx    = int(elo_ser.loc[best_b_idx]   .idxmin()) if not best_b_idx    .empty else None
                        worst_idx   = int(elo_ser.loc[worst_b_idx]  .idxmax()) if not worst_b_idx   .empty else None

                    stats_hl[col] = {'best_idx': best_idx, 'worst_idx': worst_idx}

        headers         = list(df_players.columns)
        html_rows_list  = []

        for _, row in df_players.iterrows():
            row_dict = {}

            for col in headers:
                val = row[col]

                if      col in ["Player", "Team", "Tier"] or isinstance(val, dict)      : row_dict[col] = val
                elif    pd.isnull(val) or (isinstance(val, float) and np.isnan(val))    : row_dict[col] = "N/A"
                elif    col in int_cols                                                 : row_dict[col] = int(val)
                else                                                                    : row_dict[col] = f"{float(val):.2f}"

            html_rows_list.append(row_dict)

        return html_rows_list, stats_hl, borders, eligibility

    def _render_dashboard_tour(self, watched, tour_song_details, player_song_details):
        stats           = self._compute_tour_stats(self.use_teams, watched)
        tour_unrolled   = []

        for row in stats:
            metric_name = row[0]
            display_val = str(row[1])
            link_key    = row[2]
            details     = []

            if link_key is not None:
                if isinstance(link_key, str): details = tour_song_details.get(link_key, [])

                elif isinstance(link_key, tuple):
                    stat_key, player_name   = link_key
                    lookup_key              = "Solo Rig Conversions" if "Converter" in metric_name else stat_key
                    details                 = player_song_details.get(player_name, {}).get(lookup_key, [])
            
            details.sort(key = str.lower)
            tour_unrolled.append({"Metric": metric_name, "Value": {"count": display_val, "details": details}})

        return tour_unrolled

    def _render_dashboard_team(self, df_teams, team_song_details):
        team_rows, team_hl_rules = [], {}
        if not self.use_teams: return team_rows, team_hl_rules

        for _, row_data in df_teams.iterrows():
            tid = row_data["_tid"]
            h   = row_data["_history"]

            team_song_details[tid]["Total 1/8s"].sort(key = str.lower)

            details_hover = []

            if h["total_matches"] > 0:
                def line_fmt(header, sub_dict):
                    if not sub_dict: return ""
                    items = []
                    for opp, scores in sub_dict.items(): items.append(f"{opp} ({', '.join(scores)})")
                    return f"{header}: {', '.join(items)}"

                w_line = line_fmt("Win",    h["wins"])
                l_line = line_fmt("Loss",   h["losses"])
                t_line = line_fmt("Tie",    h["ties"])

                if w_line: details_hover.append(w_line)
                if l_line: details_hover.append(l_line)
                if t_line: details_hover.append(t_line)

            summary_parts   = [int(x) for x in h["summary"].split("-")]
            tot_m           = sum(summary_parts)
            tie_val         = summary_parts[2] if len(summary_parts) > 2 else 0
            win_rate_val    = ((summary_parts[0] + 0.5 * tie_val) / tot_m * 100) if tot_m > 0 else np.nan

            item_payload = {
                "Team Leader"   : row_data["Team Leader"],
                "Mean Elo"      : float (row_data["Mean Elo"]),
                "Mean GR"       : float (row_data["Mean GR"]),
                "Total 1/8s"    : {"count": int(row_data["Total 1/8s"]), "details": team_song_details[tid]["Total 1/8s"]},
                "Mean Over-8"   : float (row_data["Mean Over-8"]),
                "Rig Synergy"   : float (row_data["Rig Synergy"]),
                "Off Synergy"   : float (row_data["Off Synergy"]),
                "Shared Rigs"   : float (row_data["Shared Rigs"])
            }

            if h["total_matches"] > 0:
                item_payload["Win Record"]      = {"count": h["summary"], "details": details_hover}
                item_payload["_win_pct_sort"]   = win_rate_val

            else: item_payload["_win_pct_sort"] = -1.0

            team_rows.append(item_payload)

        if team_rows:
            df_teams_temp = pd.DataFrame(team_rows)

            for col in df_teams_temp.columns:
                if col.startswith("_"): continue

                num     = df_teams_temp[col].map(lambda x: x["count"] if isinstance(x, dict) else x)
                desc    = ["Mean Elo", "Mean GR", "Total 1/8s", "Rig Synergy", "Off Synergy", "Shared Rigs", "Win Record"]
                asc     = ["Mean Over-8"]

                if not num.dropna().empty and (col in desc or col in asc):
                    clean_num = num.loc[df_teams_temp["_win_pct_sort"] != -1.0] if col == "Win Record" else num
                    if clean_num.dropna().empty: continue

                    best_val    = clean_num.dropna().min() if col in asc else clean_num.dropna().max()
                    worst_val   = clean_num.dropna().max() if col in asc else clean_num.dropna().min()

                    best_b_idx  = num[num == best_val]  .index
                    worst_b_idx = num[num == worst_val] .index

                    team_hl_rules[col] = {
                        'best_idx'  : int(best_b_idx    [0]) if not best_b_idx  .empty else None,
                        'worst_idx' : int(worst_b_idx   [0]) if not worst_b_idx .empty else None
                    }

        formatted_team_rows = []

        for row in team_rows:
            f_dict = {}

            for k, v in row.items():
                if      k.startswith("_")                                       : continue
                if      k in ["Total 1/8s", "Win Record", "Team Leader"]        : f_dict[k] = v
                elif    pd.isnull(v) or (isinstance(v, float) and np.isnan(v))  : f_dict[k] = "N/A"
                else                                                            : f_dict[k] = f"{float(v):.2f}"

            formatted_team_rows.append(f_dict)

        return formatted_team_rows, team_hl_rules

    def _render_dashboard_tier(self, rows1, rows2, player_song_details):
        tier_data = {}

        for r1, r2 in zip(rows1, rows2):
            tr              = r1["Tier"]
            tier_data[tr]   = []
            players_tracked = {p["player"] for p in r1["_players"]["gen"]}

            for p in players_tracked:
                tot = self.s_part   [p]
                cor = self.c_counts [p]
                chc = self.p_chan_c [p]
                cht = self.p_chan_s [p]

                gen = 100 * cor / tot if tot else 0.0
                atk = next((x["value"] for x in r1["_players"]["atk"] if x["player"] == p), 0.0)
                blk = next((x["value"] for x in r1["_players"]["blk"] if x["player"] == p), 0.0)
                con = 100 * (atk + blk) / cor if cor else 0.0
                spd = next((x["value"] for x in r2["_players"]["spd"] if x["player"] == p), None)
                chn = 100 * chc / cht if cht else 0.0

                player_song_details[p]["Lives Taken"].sort(key = str.lower)
                player_song_details[p]["Lives Saved"].sort(key = str.lower)

                taken_suffix = {}
                saved_suffix = {}

                for s in player_song_details[p]["Lives Taken"]:
                    parts               = s.split(" (from", 1)
                    song                = parts[0].strip().lower()
                    taken_suffix[song]  = f"(taken from{parts[1]}" if len(parts) > 1 else ""

                for s in player_song_details[p]["Lives Saved"]:
                    parts               = s.split(" (from", 1)
                    song                = parts[0].strip().lower()
                    saved_suffix[song]  = f"(saved from{parts[1]}" if len(parts) > 1 else ""

                contribution_details = []

                for song_line in player_song_details[p]["Overall"]:
                    if not song_line.startswith("✓"): continue
                    song = song_line[2:].strip().lower()

                    if      song in taken_suffix    : contribution_details.append(f"✓ {song_line[2:]} {taken_suffix[song]}")
                    elif    song in saved_suffix    : contribution_details.append(f"✓ {song_line[2:]} {saved_suffix[song]}")
                    else                            : contribution_details.append(f"✗ {song_line[2:]}")

                times = self.p_answer_times.get(p, [])

                if times and spd is not None and pd.notnull(spd):
                    t_min   = float(np.min(times))
                    t_mean  = float(np.mean(times))
                    t_max   = float(np.max(times))
                    t_std   = float(np.std(times))
                    t_det   = {
                        "count"     : float(round(spd, 2)),
                        "details"   : [
                            f"Minimum: {t_min:.2f}",
                            f"Mean: {t_mean:.2f}",
                            f"Maximum: {t_max:.2f}",
                            f"Standard Deviation: {t_std:.2f}"
                        ]
                    }

                else: t_det = None

                tier_data[tr].append({
                    "Player"                : p,
                    "GR"                    : {"count": float(round(gen, 2)), "details": [f"{cor}/{tot}"] + player_song_details[p]["Overall"]},
                    "Lives Taken"           : int(atk),
                    "Lives Taken Details"   : player_song_details[p]["Lives Taken"],
                    "Lives Saved"           : float(round(blk, 2)),
                    "Lives Saved Details"   : player_song_details[p]["Lives Saved"],
                    "Contribution Rate"     : {"count": float(round(con, 2)), "details": [f"{int(atk) + int(blk)}/{cor}"] + contribution_details},
                    "Median Time"           : t_det,
                    "Chanting GR"           : {"count": float(round(chn, 2)), "details": [f"{chc}/{cht}"] + player_song_details[p]["Chant"]}
                })

        return tier_data

    def _render_dashboard_song(self):
        song_matrix_list = []

        for s in self.song_data:
            if s["vintage"] > 0: song_matrix_list.append({
                "vintage"       : float(round(s["vintage"],     2)), 
                "difficulty"    : float(round(s["difficulty"],  2)), 
                "correct_count" : int(s["correct_count"])
            })

        return song_matrix_list

    def _render_dashboard_plot(self, avg_rank, raw_vintage_by_guess, raw_vintage_by_list):
        pool_data = []

        for name in self.s_part:
            if self.c_counts[name] > 0:
                tot         = self.s_part[name]
                uf_scaled   = (self.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0

                try     : elo = float(self.elo_map.get(name.lower(), 0.0))
                except  : elo = 0.0

                pool_data.append({"name": name, "uf": uf_scaled, "elo": elo})

        els = np.array([p["elo"]    for p in pool_data])
        ufs = np.array([p["uf"]     for p in pool_data])

        if len(els) > 1 and np.var(els) > 0:
            slope, intercept    = np.polyfit(els, ufs, 1)
            res_std             = np.std(ufs - (slope * els + intercept))

            if res_std == 0: res_std = 1

        else: slope, intercept, res_std = 0, np.mean(ufs) if len(ufs) > 0 else 0, 1

        scatter_list, arrow_list = [], []

        for name in self.s_part:
            if self.c_counts[name] > 0:
                yl = np.median(self.p_l_vint[name]) if self.p_l_vint[name] else np.nan
                yg = np.median(self.p_c_vint[name]) if self.p_c_vint[name] else np.nan
                
                p_vints     = raw_vintage_by_guess.get(name, [])
                p_vint_med  = np.median([extract_year(v) for v in p_vints]) if p_vints else yg
                p_seas      = format_year(p_vint_med)                       if p_vints else f"Winter {int(yg)}" if pd.notnull(yg) else "N/A"
                
                r_vints     = raw_vintage_by_list.get(name, [])
                r_vint_med  = np.median([extract_year(v) for v in r_vints]) if r_vints else yl
                r_seas      = format_year(r_vint_med)                       if r_vints else f"Winter {int(yl)}" if pd.notnull(yl) else "N/A"

                tot         = self.s_part[name]
                uf_scaled   = (self.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0

                try     : elo = float(self.elo_map.get(name.lower(), 0.0))
                except  : elo = 0.0
                
                expected_uf = slope * elo + intercept
                residual    = uf_scaled - expected_uf
                perf_score  = (1 / (1 + np.exp(SCALE_PERF * (residual / res_std)))) * 100

                base_node = {
                    "acronym"           : self._get_player_acronym(name),
                    "name"              : name,
                    "over8"             : float(round(self.p_overs_sum  [name] / self.c_counts[name],       2)),
                    "vintage"           : float(round(p_vint_med,                                           2)),
                    "seasonal_vintage"  : p_seas,
                    "gr"                : float(round(self.c_counts     [name] / self.s_part[name] * 100,   2)) if self.s_part[name] else 0.0,
                    "rig_gr"            : float(round(self.p_rigs_h     [name] / self.p_rigs[name] * 100,   2)) if self.p_rigs[name] else 0.0,
                    "performance"       : float(round(perf_score,                                           2)),
                    "rig_rate"          : float(round(self.p_rigs       [name] / self.s_part[name] * 100,   2)) if self.s_part[name] else 0.0
                }

                scatter_list.append(base_node)

                if self.p_l_corr[name] and pd.notnull(yl) and pd.notnull(yg): 
                    hit_over8   = np.mean   (self.p_lh_corr[name]) if self.p_lh_corr[name] else base_node["over8"]
                    hit_vint    = np.median (self.p_lh_vint[name]) if self.p_lh_vint[name] else base_node["vintage"]

                    arrow_list.append({
                        "acronym"                   : base_node["acronym"],
                        "name"                      : name,
                        "x_start"                   : float(round(np.mean(self.p_l_corr[name]), 2)),
                        "y_start"                   : float(round(r_vint_med,                   2)),
                        "seasonal_vintage_start"    : r_seas,
                        "x_end"                     : base_node["over8"],
                        "y_end"                     : base_node["vintage"],
                        "x_hit"                     : float(round(hit_over8,                    2)),
                        "y_hit"                     : float(round(hit_vint,                     2)),
                        "seasonal_vintage_end"      : p_seas,
                        "rig_gr"                    : base_node["rig_gr"],
                        "gr"                        : base_node["gr"],
                        "rig_rate"                  : base_node["rig_rate"]
                    })

        return scatter_list, arrow_list

    def _create_dashboard_html(self, path, use_teams, watched):
        active = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]
        if len(active) <= 1: active = []

        t_labels        = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
        valid_elos      = [float(v) for v in self.elo_map.values() if str(v).replace('.', '', 1).isdigit() or (str(v).startswith('-') and str(v)[1:].replace('.', '', 1).isdigit())]
        avg_rank        = np.mean(valid_elos) if valid_elos else 1.0
        final_threshold = 6 if len(self.s_part) <= THRESH_PLYR else 5

        if      self.base_exp >= final_threshold    : stage = "Final"
        elif    self.base_exp == 3                  : stage = "Mid-Tour"
        else                                        : stage = f"R{self.base_exp}"

        prefix = f"{self.tour_label.strip()} Tour: {stage}"
        player_song_details, tour_song_details, team_song_details, raw_vintage_by_guess, raw_vintage_by_list, matrix_song_details, num_x, num_y = self._get_dashboard_data()
        df_base, _ = self._compute_player_rows(self.elo_map, self.apps, self.exp_map, self.base_exp, self.new_players, watched, active, t_labels, avg_rank)

        def player_sort_key(x):
            gr = (self.c_counts[x] / self.s_part[x]) if self.s_part[x] else 0.0

            try     : elo = float(self.elo_map.get(x.lower(), float('inf')))
            except  : elo = float('inf')

            return (gr, -elo)

        sorted_players  = sorted(self.s_part.keys(), key = player_sort_key, reverse = True)
        df_teams        = self._compute_team_rows(self.assignments, self.t1_lookup)
        rows1, rows2    = self._compute_tier_rows(self.assignments, any(self.p_chan_s.values()))

        render_players, render_hl_rules, render_borders, render_eligibility = self._render_dashboard_player(sorted_players, active, t_labels, watched, player_song_details, df_base)
        render_teams, render_team_hl_rules                                  = self._render_dashboard_team(df_teams, team_song_details)
        render_scatter, render_arrows                                       = self._render_dashboard_plot(avg_rank, raw_vintage_by_guess, raw_vintage_by_list)

        render_tour_stats     = self._render_dashboard_tour(watched, tour_song_details, player_song_details)
        render_tier_merged    = self._render_dashboard_tier(rows1, rows2, player_song_details)
        render_songs          = self._render_dashboard_song()

        explanations = {
            "Player"                    : "★ New player<br>▲ Subbed in<br>▼ Subbed out",
            "UF"                        : "Usefulness<br>Calculates this player's contribution to their team, scaled by Elo and songs played",
            "Score"                     : "Calculates this player's value (Usefulness) against what's expected from their Elo<br>50 means this player is playing to expectations",
            "Mean Over-8"               : "Average of correct guessers across songs this player/team guessed correctly",
            "Lives Taken"               : "Count of points won against the opposing team<br>Correct guessers exclusively on their team",
            "Lives Saved"               : "Count of blocks achieved against the opposing team<br>Lone correct guesser for their team whilst the opposing team also has correct guesser(s)",
            "Solo Rigs"                 : "Count of songs exclusively from this player's list",
            "Rig Over-8"                : "Average of correct guessers across songs from this player's list",
            "Over-8 Δ"                  : "Rig Over-8 - Mean Over-8<br>Calculates the difficulty gap between this player's list and correct guesses",
            "Rig Δ"                     : "100 * (Correct - Rig) / Correct<br>Calculates this player's performance against their own list",
            "Median Time"               : "Median guess time across songs this player guessed correctly",
            "Total 4-0s"                : "Count of songs where all players from one team guessed correctly and all players from the other team missed",
            "Rig Synergy"               : "Average team guess rate across songs from its own members' lists",
            "Off Synergy"               : "Average team guess rate across songs from the opposing team member's lists",
            "Shared Rigs"               : "Calculates how much songs are shared across its own members' lists",
            "Contribution Rate"         : "100 * (Lives Taken + Saved) / Correct<br>Calculates how much of this player's correct guesses directly contributed to the scoreline",
            "Best Solo Rig Converter"   : "100 * Solo from Solo Rig / Solo Rig<br>Shows the best player at converting their own solo rig into a solo",
            "Worst Solo Rig Converter"  : "100 * Solo from Solo Rig / Solo Rig<br>Shows the worst player at converting their own solo rig into a solo"
        }

        data_payload = {
            "prefix"                : prefix,
            "use_teams"             : use_teams,
            "watched"               : watched,
            "num_x"                 : num_x,
            "num_y"                 : num_y,
            "c0"                    : COLOR_0,
            "c1"                    : COLOR_1,
            "c2"                    : COLOR_2,
            "json_players"          : render_players,
            "json_hl_rules"         : render_hl_rules,
            "json_borders"          : render_borders,
            "json_eligibility"      : render_eligibility,
            "json_tour_stats"       : render_tour_stats,
            "json_teams"            : render_teams,
            "json_team_hl_rules"    : render_team_hl_rules,
            "json_tier_merged"      : render_tier_merged,
            "json_songs"            : render_songs,
            "json_matrix_songs"     : matrix_song_details,
            "json_scatter"          : render_scatter,
            "json_arrows"           : render_arrows,
            "json_explanations"     : explanations,
            "generated_timestamp"   : int(time.time() * 1000)
        }

        search_songs_list   = []

        for path_json in self.json_paths:
            try:
                with open(path_json, encoding = "utf-8") as f: data_j = json.load(f)

            except: continue

            raw_f_players = set()

            for s in data_j.get("songs", []):
                for p in s.get("correctGuessPlayers", []):
                    if      isinstance(p, str)                  : raw_f_players.add(p)
                    elif    isinstance(p, dict) and "name" in p : raw_f_players.add(p["name"])

                for ls in s.get("listStates", []):
                    if "name" in ls: raw_f_players.add(ls["name"])

            final_members = set(raw_f_players)

            if self.use_teams:
                t_in_f = {self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments}

                for tid in t_in_f:
                    ros = self.rosters[tid]

                    for m_p in ros:
                        if m_p.lower() not in self.assignments:
                            for c_p in raw_f_players:
                                if c_p.lower() in self.assignments and self.assignments[c_p.lower()][0] == tid: self.assignments[m_p.lower()] = self.assignments[c_p.lower()]

                if len(final_members) < 8:
                    for tid in t_in_f: final_members.update(self.rosters[tid])

            room_players_list = sorted(list(final_members))

            for song in data_j.get("songs", []):
                si          = song.get("songInfo", {})
                ann_id_raw  = si.get("annId")

                if not ann_id_raw: continue

                lives_taken_players = []
                lives_saved_players = []

                anime_romaji    = si.get("animeNames",  {}).get("romaji",  "Unknown").strip()
                anime_english   = si.get("animeNames",  {}).get("english", "").strip()
                song_name       = si.get("songName",    "Unknown").strip()
                raw_artist      = si.get("artist",      "Unknown").strip()
                artist_arr      = [a.strip() for a in raw_artist.split(",") if a.strip()] if raw_artist else []
                composer_name   = si.get("composerInfo", {}).get("name", "Unknown").strip() if si.get("composerInfo") else "Unknown"
                arranger_name   = si.get("arrangerInfo", {}).get("name", "Unknown").strip() if si.get("arrangerInfo") else "Unknown"

                st     = si.get("type",         3)
                t_num  = si.get("typeNumber",   0)

                if   st == 1 : type_fmt = f"Opening {t_num}"
                elif st == 2 : type_fmt = f"Ending {t_num}"
                else         : type_fmt = "Insert"

                ann_url         = f"https://www.animenewsnetwork.com/encyclopedia/anime.php?id={str(ann_id_raw)}"
                anime_type_raw  = str(si.get("animeType", "N/A")).strip()

                if      anime_type_raw.lower() == "movie"   : anime_type = "Movie"
                elif    anime_type_raw.lower() == "special" : anime_type = "Special"
                else                                        : anime_type = anime_type_raw

                vint_raw = str(si.get("vintage", "Unknown")).strip().replace("\n", " ").replace("\r", " ")

                try:
                    diff_val    = si.get("animeDifficulty")
                    safe_diff   = f"{float(diff_val):.2f}" if diff_val is not None and float(diff_val) > 0 else "Unrated"

                except: safe_diff = "Unrated"

                video_url       = song.get("videoUrl",              "")
                raw_correct     = song.get("correctGuessPlayers",   [])
                ann_song_id_str = str(si.get("annSongId",           ""))

                is_chanting_str = "Yes" if ann_song_id_str in self.chanting_ids else "No"
                guess_times     = {}

                for p in raw_correct:
                    if isinstance(p, dict) and "name" in p:
                        t_val = p.get("answerTime")
                        guess_times[p["name"].lower()] = f"{float(t_val):.2f}" if t_val is not None else "N/A"

                    elif isinstance(p, str): guess_times[p.lower()] = "N/A"

                guessers_flat = []

                for p in raw_correct:
                    if      isinstance(p, str)                  : guessers_flat.append(p)
                    elif    isinstance(p, dict) and "name" in p : guessers_flat.append(p["name"])

                raw_lists       = song.get("listStates", [])
                listers_flat    = [ls["name"] for ls in raw_lists if isinstance(ls, dict) and "name" in ls]

                if self.use_teams:
                    t_list = list({self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments})

                    if len(t_list) == 2:
                        tA, tB      = t_list[0], t_list[1]
                        correct_set = set(guessers_flat)
                        cA, cB      = correct_set & self.rosters[tA], correct_set & self.rosters[tB]

                        if len(cA) == 4 and not cB: self.t_sweeps[tA] += 1; self.global_stats["sweeps"] += 1
                        if len(cB) == 4 and not cA: self.t_sweeps[tB] += 1; self.global_stats["sweeps"] += 1

                        for cur, opp in [(tA, tB), (tB, tA)]:
                            cC, oC = correct_set & self.rosters[cur], correct_set & self.rosters[opp]

                            if not oC: 
                                for p in cC: 
                                    lives_taken_players.append(p.lower())

                            if len(cC) == 1 and len(oC) > 0: 
                                lone_p = list(cC)[0]
                                lives_saved_players.append(lone_p.lower())

                def group_by_team_structure(target_players, include_times = False):
                    if not self.use_teams:
                        sorted_p = sorted(target_players, key = str.lower)
                        if include_times: return [f"{p} ({guess_times.get(p.lower(), 'N/A')})" for p in sorted_p]
                        return sorted_p

                    team_buckets = defaultdict(list)

                    for p in target_players:
                        tid, _ = self.assignments.get(p.lower(), (None, "5"))
                        team_buckets[tid].append(p)

                    sorted_tids = sorted([t for t in team_buckets.keys() if t is not None])
                    hover_lines = []

                    for tid in sorted_tids:
                        leader_name = self.t1_lookup.get(tid, f"Team {tid}")
                        pts_sorted  = sorted(team_buckets[tid], key=lambda x: self.assignments.get(x.lower(), (None, "5"))[1])
                        p_strings   = []

                        for p in pts_sorted:
                            if include_times    : p_strings.append(f"{p} ({guess_times.get(p.lower(), 'N/A')})")
                            else                : p_strings.append(p)

                        hover_lines.append(f"Team {leader_name}: {', '.join(p_strings)}")

                    all_active_tids = {self.assignments[p.lower()][0] for p in final_members if p.lower() in self.assignments}
                    missing_tids    = sorted(list(all_active_tids - set(team_buckets.keys())))

                    for tid in missing_tids:
                        leader_name = self.t1_lookup.get(tid, f"Team {tid}")
                        hover_lines.append(f"Team {leader_name}: None")

                    return hover_lines

                guessers_hover  = group_by_team_structure(guessers_flat,    include_times = True)
                listers_hover   = group_by_team_structure(listers_flat,     include_times = False)

                raw_genres      = si.get("animeGenre",              [])                             if isinstance(si.get("animeGenre"),             list) else []
                raw_tags        = [t for t in si.get("animeTags",   []) if t not in EXCLUDED_TAGS]  if isinstance(si.get("animeTags"),              list) else []
                raw_alts        = si.get("altAnimeNames",           [])                             if isinstance(si.get("altAnimeNames"),          list) else []
                raw_alts_ans    = si.get("altAnimeNamesAnswers",    [])                             if isinstance(si.get("altAnimeNamesAnswers"),   list) else []
                titles_to_check = {anime_romaji.lower().strip(), anime_english.lower().strip()}
                combined_alts   = list({alt.strip() for alt in (raw_alts + raw_alts_ans) if isinstance(alt, str) and alt.strip().lower() not in titles_to_check})

                search_songs_list.append({
                    "romaji"                : anime_romaji,
                    "english"               : anime_english,
                    "song"                  : song_name,
                    "artist_raw"            : raw_artist,
                    "artist_arr"            : artist_arr,
                    "composer"              : composer_name,
                    "arranger"              : arranger_name,
                    "type"                  : type_fmt,
                    "chanting"              : is_chanting_str,
                    "ann_url"               : ann_url,
                    "anime_type"            : anime_type,
                    "vintage"               : vint_raw,
                    "genres_raw"            : raw_genres,
                    "tags_raw"              : raw_tags,
                    "difficulty"            : safe_diff,
                    "video_url"             : video_url,
                    "guessers_flat"         : guessers_flat,
                    "guessers_hover"        : guessers_hover,
                    "listers_flat"          : listers_flat,
                    "listers_hover"         : listers_hover,
                    "room_players"          : room_players_list,
                    "lives_taken_flat"      : lives_taken_players,
                    "lives_saved_flat"      : lives_saved_players,
                    "correct_teams_flat"    : [self.assignments[p.lower()][0] for p in guessers_flat if p.lower() in self.assignments],
                    "alts"                  : combined_alts,
                    "start_sample"          : song.get("startPoint", 0)
                })

        search_songs_list.sort(key = lambda x: x["romaji"].lower())

        with open(path / "Search.json", "w", encoding = "utf-8") as f: json.dump(search_songs_list, f, ensure_ascii = False, indent = 4)
        with open(path / "Data.json",   "w", encoding = "utf-8") as f: json.dump(data_payload,      f, ensure_ascii = False, indent = 4)

        template_dir = self.script_dir / "help" / "template"

        shutil.copy(template_dir / "index.html",    path / "index.html")
        shutil.copy(template_dir / "Styles.css",    path / "Styles.css")
        shutil.copy(template_dir / "Script.js",     path / "Script.js")
        shutil.copy(template_dir / "Name.json",     path / "Name.json")

    def _export_png(self, df, path, fname, title, mask = None, val_str = "default"):
        if not self.browser_path: return
        df = df.reset_index(drop = True)

        delta_check_cols    = ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]
        cols_to_drop        = [c for c in delta_check_cols if c in df.columns and (df[c].isna() | (df[c].astype(str).str.strip() == "N/A")).all()]

        if cols_to_drop: df = df.drop(columns = cols_to_drop)

        desc = [
            "Elo",
            "GR",
            "GR Δ",
            "UF",
            "UF Δ",
            "Score",
            "1/8s",
            "2/8s",
            "Lives Taken",
            "Lives Saved",
            "OP GR",
            "OP Δ",
            "ED GR",
            "ED Δ",
            "IN GR",
            "IN Δ",
            "Rigs",
            "Rig Rate",
            "Solo Rigs",
            "Solo Rig Rate",
            "Over-8 Δ",
            "Rig GR",
            "Off GR",
            "Rig Δ",
            "Median Vintage Hit",
            "Chant GR",
            "Mean Elo",
            "Mean GR",
            "Total 1/8s",
            "Rig Synergy",
            "Off Synergy",
            "Shared Rigs",
            "Win Record"
        ]

        asc     = ["7/8s", "Median Time", "Mean Over-8", "Rig Over-8", "Mean Difficulty Hit"]
        rest    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs"]
        stats   = {}
        elo_col = "Elo" if "Elo" in df.columns else "Mean Elo" if "Mean Elo" in df.columns else None
        elo_ser = pd.to_numeric(df[elo_col], errors = 'coerce').fillna(0.0) if elo_col else pd.Series(0.0, index = df.index)

        for col in df.columns:
            if col in desc or col in asc:
                if col == "Win Record":
                    def parse_wlt(val):
                        try:
                            parts = [int(x) for x in str(val).split("-")]
                            total = sum(parts)
                            ties  = parts[2] if len(parts) > 2 else 0

                            return ((parts[0] + 0.5 * ties) / total) if total > 0 else -1.0

                        except: return -1.0

                    num = df[col].apply(parse_wlt)

                else: num = pd.to_numeric(df[col].astype(str).str.replace('%',''), errors = 'coerce')

                el_num = num[mask].dropna() if mask is not None and col in rest else num.dropna()

                if not num.dropna().empty:
                    if col in desc:
                        best_val    = num.dropna().max()
                        worst_val   = el_num.min() if not el_num.empty else None

                    else:
                        best_val    = num.dropna().min()

                        if col == "Median Time":
                            under_limit = el_num[el_num < THRESH_TIME]
                            worst_val   = under_limit.max() if not under_limit.empty else None

                        else: worst_val = el_num.max() if not el_num.empty else None

                    best_b_indices  = num       [num    == best_val]    .index if pd.notnull(best_val)  else pd.Index([])
                    worst_b_indices = el_num    [el_num == worst_val]   .index if pd.notnull(worst_val) else pd.Index([])

                    el_cols = ["Elo", "Mean Elo"]
                    gr_cols = ["OP GR", "OP Δ", "ED GR", "ED Δ", "IN GR", "IN Δ", "Chant GR"]
                    rig_ser = pd.to_numeric(df["Rigs"], errors = 'coerce').fillna(0) if "Rigs" in df.columns else pd.Series(0, index = df.index)

                    if col == "Solo Rigs":
                        best_idx    = rig_ser.loc[best_b_indices]   .idxmin() if not best_b_indices     .empty else None
                        worst_idx   = rig_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                    elif col == "Solo Rig Rate":
                        best_idx    = rig_ser.loc[best_b_indices]   .idxmax() if not best_b_indices     .empty else None
                        worst_idx   = rig_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                    elif col in el_cols:
                        best_idx    = best_b_indices    [0] if not best_b_indices   .empty else None
                        worst_idx   = worst_b_indices   [0] if not worst_b_indices  .empty else None

                    elif col in gr_cols:
                        best_idx    = pd.to_numeric(df["GR"], errors = 'coerce').fillna(0).loc[best_b_indices]  .idxmin() if not best_b_indices     .empty else None
                        worst_idx   = pd.to_numeric(df["GR"], errors = 'coerce').fillna(0).loc[worst_b_indices] .idxmax() if not worst_b_indices    .empty else None

                    elif col == "Rig GR" and "Rigs" in df.columns:
                        best_idx    = rig_ser.loc[best_b_indices]   .idxmax() if not best_b_indices     .empty else None
                        worst_idx   = elo_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                    else:
                        best_idx    = elo_ser.loc[best_b_indices]   .idxmin() if not best_b_indices     .empty else None
                        worst_idx   = elo_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                    stats[col] = {'best_idx': best_idx, 'worst_idx': worst_idx}

        borders = []

        if "GR" in df.columns:
            if "Eru" in self.tour_label: th = []

            else:
                if val_str == "default":
                    if      self.tour_label == "Watched 2+8s"               : th_val = "25, 20, 15, 10, 5"
                    elif    self.tour_label in ["Watched", "QuagWatched"]   : th_val = "28, 18, 12, 6"
                    elif    self.tour_label in ["Usual", "Quagsual"]        : th_val = "28, 19, 8"
                    elif    "Rigs"          in df.columns                   : th_val = "28, 18, 12, 6"
                    else                                                    : th_val = "28, 19, 8"

                else: th_val = val_str

                try     : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
                except  : th = [28.0, 18.0, 12.0, 6.0]

            gv = pd.to_numeric(df["GR"].astype(str).str.replace('%',''), errors = 'coerce').tolist()

            for t in th:
                f_idx = -1

                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i

                if f_idx != -1 and f_idx < len(df) - 1: borders.append(f_idx)

        col_borders = {
            "Player",
            "Score",
            "Mean Over-8",
            "Lives Saved",
            "IN Δ",
            "Rig Rate",
            "Solo Rig Rate",
            "Over-8 Δ",
            "Rig Δ",
            "Median Vintage Hit",
            "Metric",
            "Value",
            "Team Leader"
        }

        if "Score"  not in df.columns: col_borders.add("GR Δ") if "GR Δ" in df.columns else col_borders.add("GR")
        if "IN Δ"   not in df.columns: col_borders.add("IN GR")

        th_cells = []

        for c in df.columns:
            s_th = ' style="border-right: 3px solid black;"' if c in col_borders else ''
            th_cells.append(f"<th{s_th}>{str(c).replace(' ', '<br>')}</th>")

        html            = "<thead><tr>" + "".join(th_cells) + "</tr></thead><tbody>"
        bold_columns    = {"Player", "Metric", "Team Leader"}

        for idx, row in df.iterrows():
            b_s     =   "border-bottom: 3px solid black;" if idx in borders else ""
            html    +=  "<tr>"

            for i, (cname, cell) in enumerate(row.items()):
                style = [b_s] if b_s else []
                if cname in col_borders: style.append("border-right: 3px solid black;")

                if      cname == "Mean Difficulty Hit"  and pd.notnull(cell) and isinstance(cell, (int, float)) : cell_display = f"{float(cell):.2f}"
                elif    cname == "Median Vintage Hit"   and pd.notnull(cell) and isinstance(cell, (int, float)) : cell_display = format_year(float(cell))
                else                                                                                            : cell_display = cell

                if cname in stats:
                    val_best_idx    = stats[cname]['best_idx']
                    val_worst_idx   = stats[cname]['worst_idx']

                    if cname in desc:
                        is_max = (idx == val_best_idx)
                        is_min = (idx == val_worst_idx)

                    else:
                        is_max = (idx == val_worst_idx)
                        is_min = (idx == val_best_idx)

                    elig = True if mask is None or cname not in rest else mask[idx]

                    if cname in desc:
                        if      is_max          : style.append(f"background-color: {COLOR_2}; color: white; font-weight: bold;")
                        elif    is_min and elig : style.append(f"background-color: {COLOR_0}; color: white; font-weight: bold;")

                    elif cname in asc:
                        if      is_max and elig : style.append(f"background-color: {COLOR_0}; color: white; font-weight: bold;")
                        elif    is_min          : style.append(f"background-color: {COLOR_2}; color: white; font-weight: bold;")

                s_attr  =   f' style="{" ".join(style)}"' if style else ""
                cnt     =   f"<b>{cell}</b>" if cname in bold_columns else cell_display
                html    +=  f"<td{s_attr}>{cnt}</td>"

            html += "</tr>"

        full = f"""<html>
            <head>
                <style>
                    body {{
                        font-family         : 'Segoe UI';
                        background          : white;
                        display             : inline-block;
                        margin              : 0
                    }} 
                    h2 {{
                        margin              : 0 0 10px 0;
                        font-size           : 40px;
                        text-align          : center
                    }} 
                    table {{
                        border-collapse     : collapse;
                        width               : auto;
                        border              : 3px solid black
                    }} 
                    th {{
                        font-weight         : bold;
                        font-size           : 25px;
                        text-align          : center;
                        padding             : 10px;
                        border              : 1px solid black;
                        border-bottom       : 3px solid black;
                        background-color    : #f0f0f0
                    }} 
                    td {{
                        font-size           : 25px;
                        text-align          : center;
                        padding             : 10px;
                        border              : 1px solid black
                    }}
                    tr:nth-child(even) {{background-color: #f0f0f0}}
                </style>
            </head>
            <body>
                <h2>{title}</h2>
                <table>{html}</table>
            </body>
        </html>"""

        hti = Html2Image(size = (max(2000, len(df.columns) * 120), max(2000, len(df) * 60)), browser_executable = self.browser_path, output_path = str(path), custom_flags = ['--log-level=3', '--silent'])
        hti.screenshot(html_str = full, save_as = fname)

        try     : trim_whitespace(path / fname)
        except  : pass

    def _fuse(self, path):
        f       = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "Guess": "Guess.png", "List": "List.png", "Song": "Song.png"}
        ps      = {k: path / v for k, v in f.items() if (path / v).exists()}
        imgs    = {k: Image.open(v) for k, v in ps.items()}

        if "Tour" not in imgs:
            for k, p in ps.items():
                if k not in ["List", "Guess", "Song"]:
                    try     : os.remove(p)
                    except  : pass

            return

        img_tour    = imgs.get("Tour")
        img_team    = imgs.get("Team")
        img_tier    = imgs.get("Tier")
        img_song    = imgs.get("Song")
        img_guess   = imgs.get("Guess")
        img_list    = imgs.get("List")
        img_plots   = None

        if img_song and img_guess:
            w, h                = img_song.width, img_song.height
            img_guess_scaled    = img_guess.resize((w, h), Image.Resampling.LANCZOS)

            if img_list:
                img_list_scaled     = img_list.resize((w, h), Image.Resampling.LANCZOS)
                plots_w, plots_h    = w, h + 10 + h + 10 + h
                img_plots           = Image.new("RGB", (plots_w, plots_h), "white")

                img_plots.paste(img_song,           (0, 0))
                img_plots.paste(img_guess_scaled,   (0, h + 10))
                img_plots.paste(img_list_scaled,    (0, h + 10 + h + 10))

            else:
                plots_w, plots_h    = w, h + 10 + h
                img_plots           = Image.new("RGB", (plots_w, plots_h), "white")

                img_plots.paste(img_song,           (0, 0))
                img_plots.paste(img_guess_scaled,   (0, h + 10))

        if img_plots:
            plots_out_p = path / "Plots.png"
            img_plots.save(plots_out_p, compress_level = 9, optimize = True)

            try     : trim_whitespace(plots_out_p)
            except  : pass

            img_plots = Image.open(plots_out_p)

        left_components = [img_tour]

        if img_team: left_components.append(img_team)
        if img_tier: left_components.append(img_tier)

        gap_size    = 10
        total_gaps  = gap_size * (len(left_components) - 1)
        left_h_raw  = sum(img.height   for img in left_components) + total_gaps
        left_w_max  = max(img.width    for img in left_components)

        if img_plots:
            plots_aspect        = img_plots.width / img_plots.height
            plots_h_scaled      = left_h_raw
            plots_w_scaled      = int(plots_h_scaled * plots_aspect)
            img_plots_scaled    = img_plots.resize((plots_w_scaled, plots_h_scaled), Image.Resampling.LANCZOS)

        else: plots_w_scaled = 0

        w_extra     = left_w_max + (gap_size + plots_w_scaled if img_plots else 0)
        h_extra     = left_h_raw
        img_extra   = Image.new("RGB", (w_extra, h_extra), "white")
        current_y   = 0

        for img in left_components:
            centered_x = (left_w_max - img.width) // 2
            img_extra.paste(img, (centered_x, current_y))
            current_y += img.height + gap_size

        if img_plots: img_extra.paste(img_plots_scaled, (left_w_max + gap_size, 0))
        extra_out_p = path / "Extra.png"
        img_extra.save(extra_out_p, compress_level = 9, optimize = True)

        try     : trim_whitespace(extra_out_p)
        except  : pass

        img_extra = Image.open(extra_out_p)
        p_path = path / "Player.png"

        if p_path.exists():
            img_player          = Image.open(p_path)
            player_aspect       = img_player.width / img_player.height
            player_w_scaled     = img_extra.width
            player_h_scaled     = int(player_w_scaled / player_aspect)
            img_player_scaled   = img_player.resize((player_w_scaled, player_h_scaled), Image.Resampling.LANCZOS)
            gen_w               = img_extra.width
            gen_h               = player_h_scaled + 10 + img_extra.height
            img_general         = Image.new("RGB", (gen_w, gen_h), "white")

            img_general.paste(img_player_scaled, (0, 0))
            img_general.paste(img_extra, (0, player_h_scaled + 10))

            gen_out_p = path / "General.png"
            img_general.save(gen_out_p, compress_level = 9, optimize = True)

            try     : trim_whitespace(gen_out_p)
            except  : pass