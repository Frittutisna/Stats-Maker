import concurrent.futures, datetime, gspread, hashlib, json, logging, math, os, re, shutil, subprocess, sys, warnings, zipfile
import pandas   as pd
import numpy    as np

from .config        import *
from .dataloader    import *
from .dialog        import *
from .png           import *
from .web           import create_dashboard_html
from collections    import Counter, defaultdict
from curl_cffi      import requests
from pathlib        import Path
from tkinter        import messagebox

logging.getLogger("adjustText").setLevel(logging.ERROR)

warnings.filterwarnings("ignore", category  = UserWarning, module = "adjustText")
warnings.filterwarnings("ignore", message   = ".*FancyArrowPatch.*")

def _nested_int_defaultdict     (): return defaultdict(int)
def _nested_list_defaultdict    (): return defaultdict(list)

class TourAnalyzer:
    def __init__(self, tour_id: str):
        self.tour_id        = str(tour_id)
        self.script_dir     = Path(__file__).parent.parent.absolute()
        self.tour_dir       = self.script_dir / DIR_TOURS / self.tour_id
        self.browser_path   = find_browser()

        self.s_part             = defaultdict(int)
        self.c_counts           = defaultdict(int)
        self.e_counts           = defaultdict(int)
        self.p_rev_e            = defaultdict(int)
        self.p_two_e            = defaultdict(int)
        self.p_three_or_below   = defaultdict(int)
        self.p_pts              = defaultdict(int)
        self.p_blks             = defaultdict(int)
        self.p_type_c           = defaultdict(_nested_int_defaultdict)
        self.p_type_s           = defaultdict(_nested_int_defaultdict)
        self.p_rigs             = defaultdict(int)
        self.p_rigs_h           = defaultdict(int)
        self.p_l_vint           = defaultdict(list)
        self.p_c_vint           = defaultdict(list)
        self.p_l_corr           = defaultdict(list)
        self.p_lh_vint          = defaultdict(list)
        self.p_lh_corr          = defaultdict(list)
        self.p_m_erigs          = defaultdict(int)
        self.p_l_solos          = defaultdict(int)
        self.p_hit_diff         = defaultdict(list)
        self.p_hit_vint         = defaultdict(list)
        self.p_chan_c           = defaultdict(int)
        self.p_chan_s           = defaultdict(int)
        self.p_usefulness_sum   = defaultdict(float)
        self.p_overs_sum        = defaultdict(int)
        self.p_answer_times     = defaultdict(list)
        self.p_zero_e           = defaultdict(int)
        self.p_m_solos          = defaultdict(int)
        self.p_zero_x_rigs      = defaultdict(int)
        self.p_offlist_erigs    = defaultdict(int)

        self.t_vint     = defaultdict(list)
        self.t_c_ps     = defaultdict(list)
        self.t_on_syn   = defaultdict(list)
        self.t_off_syn  = defaultdict(list)
        self.t_sh_rig   = defaultdict(list)
        self.t_solos    = defaultdict(int)
        self.t_sweeps   = defaultdict(int)

        self.genre_c            = Counter()
        self.tag_c              = Counter()
        self.global_stats       = Counter()
        self.all_diff           = []
        self.all_vint           = []
        self.song_history       = []
        self.song_data          = []
        self.chanting_ids       = set()
        self.subbed_players_set = set()
        self.sub_relations      = defaultdict(list)
        self.tour_label         = ""
        self.id_database        = {}
        self.player_acronyms    = {}
        self.main_roster_names  = set()

        self.player_song_details    = defaultdict(_nested_list_defaultdict)
        self.tour_song_details      = defaultdict(list)
        self.team_song_details      = defaultdict(_nested_list_defaultdict)
        self.matrix_song_details    = defaultdict(list)
        self.raw_vintage_by_guess   = defaultdict(list)
        self.raw_vintage_by_list    = defaultdict(list)

    def prepare_configuration(self) -> bool:
        chanting_path = self.script_dir / DIR_TOURS / FILE_CHANT

        if chanting_path.exists():
            with open(chanting_path, "r", encoding = "utf-8") as f:
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
                match_digits = re.search(r"(\d+)$", path.stem)

                if match_digits:
                    m = int(match_digits.group(1))

                    if m <= THRESH_SONG:
                        songs           = songs[: min(m, len(songs))]
                        data["songs"]   = songs

            if not isinstance(songs, list) or not songs:
                messagebox.showerror("Disconnected Player JSON", f"Error in {path.name}: The exporter likely disconnected; ask someone else to re-upload this JSON")
                return False

            for song in songs:
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

        all_known, self.apps = scan_players(self.json_paths)
        self.player_acronyms = generate_acronyms(all_known)

        (
            self.use_teams,
            self.elo_map,
            self.assignments,
            self.t1_lookup,
            self.rosters,
            all_known,
            sub_candidates_raw,
            original_players_display,
        ) = load_team_data(self.tour_dir, all_known, self.id_database, self.subbed_players_set, self.main_roster_names)

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
        team_count          = len(self.rosters) if self.use_teams else 0
        baseline_initial    = int(total_jsons // (team_count / 2)) if team_count > 0 else 1
        watched_valid       = self.missing_list_count <= THRESH_WTCH

        if len(self.tour_types) == 1:
            t_map           = {1: "OP", 2: "ED", 3: "IN"}
            t_str           = t_map.get(list(self.tour_types)[0], "")
            init_label      = f"Watched {t_str}" if watched_valid else f"Random {t_str}"
        else: init_label    = "Watched" if watched_valid else "Usual"

        if "Eru" in init_label and self.use_teams   : default_th = ""
        else:
            if      init_label == "Watched 2+8s"    : default_th = "25, 20, 15, 10, 5"
            elif    "Watched" in init_label         : default_th = "28, 18, 12, 6"
            else                                    : default_th = "28, 19, 8"

        meta_dialog = TourMetadataDialog(
            None,
            self.tour_id,
            init_label,
            default_th,
            baseline_initial,
            list(all_known),
            self.elo_map,
            sub_candidates_raw,
            original_players_display,
            self.tour_dir,
        )

        if meta_dialog.result is None: sys.exit(0)

        meta_res                = meta_dialog.result
        self.mode_choice        = meta_res.get("mode_choice",       "Tour")
        self.tour_label         = meta_res["tour_label"]
        self.delta_choice       = meta_res.get("delta_choice",      "No")
        self.challonge_choice   = meta_res.get("challonge_choice",  "No")

        is_ant = (self.mode_choice == "Ant")

        stats_map = ANT_MAP_STATS               if is_ant else TOUR_MAP_STATS
        stats_key = ANT_KEY_STATS               if is_ant else TOUR_KEY_STATS
        alias_url = ANT_URL_ALIAS               if is_ant else TOUR_URL_ALIAS
        cred_name = "ant_credentials.json"      if is_ant else "credentials.json"
        auth_name = "ant_authorized_user.json"  if is_ant else "authorized_user.json"

        if self.delta_choice == "Yes" and self.tour_label in stats_map:
            print("[?] Fetching historic baselines")

            cred_file = self.script_dir / DIR_CREDS / cred_name
            auth_file = self.script_dir / DIR_CREDS / auth_name

            try:
                gc                  = gspread.oauth(credentials_filename = str(cred_file), authorized_user_filename = str(auth_file))
                sheet               = gc.open_by_key(stats_key)
                sheet_ref           = stats_map[self.tour_label]
                wks_stats           = sheet.get_worksheet_by_id(sheet_ref) if isinstance(sheet_ref, int) else sheet.worksheet(sheet_ref)
                rows_stats          = wks_stats.get_all_values()
                id_table_lookup     = load_player_ids(alias_url)
                df_stats            = pd.DataFrame(rows_stats[1:], columns = rows_stats[0])
                history_profile_map = {}
                for _, r_row in df_stats.iterrows():
                    norm_row = {str(k).strip().lower(): v for k, v in r_row.items() if pd.notnull(k)}

                    def _get_val(*keys):
                        for k in keys:
                            if k.lower() in norm_row: return norm_row[k.lower()]

                        return ""

                    def _parse_stat(val):
                        if pd.isna(val) or val == "": return 0.0

                        try:                 return float(str(val).strip().replace("%", ""))
                        except ValueError:   return 0.0

                    raw_name    = str(_get_val("player name", "name")).strip().lower()
                    pid_key     = id_table_lookup.get(raw_name)

                    if pid_key is not None: history_profile_map[pid_key] = {
                        "GR": _parse_stat(_get_val("average gr %",              "average gr",       "gr %")),
                        "UF": _parse_stat(_get_val("average usefulness (new)",  "usefulness",       "average usefulness")),
                        "OP": _parse_stat(_get_val("average ops gr %",          "average op gr %",  "op gr %")),
                        "ED": _parse_stat(_get_val("average eds gr %",          "average ed gr %",  "ed gr %")),
                        "IN": _parse_stat(_get_val("average ins gr %",          "average in gr %",  "in gr %")),
                    }

                alias_txt_path      = self.tour_dir / FILE_ALIAS
                current_alias_lines = []

                if alias_txt_path.exists():
                    with open(alias_txt_path, "r", encoding = "utf-8") as f_alias:
                        for a_line in f_alias:
                            if "," in a_line: current_alias_lines.append(a_line.strip().split(","))

                with open(alias_txt_path, "w", encoding = "utf-8") as f_out:
                    for parts in current_alias_lines:
                        p_name      = parts[0].strip()
                        alias_name  = parts[1].strip()
                        p_low       = p_name.lower()
                        p_id        = id_table_lookup.get(p_low)

                        if p_id is None and alias_name.lower() in id_table_lookup: p_id = id_table_lookup.get(alias_name.lower())

                        if p_id in history_profile_map:
                            h_prof = history_profile_map[p_id]
                            f_out.write(f"{p_name}, {alias_name}, {h_prof['GR']:.2f}, {h_prof['UF']:.2f}, {h_prof['OP']:.2f}, {h_prof['ED']:.2f}, {h_prof['IN']:.2f}\n")
                        else: f_out.write(f"{p_name}, {alias_name}, N/A, N/A, N/A, N/A, N/A\n")

                print("[✓] Historic baselines saved to alias.txt")
            except Exception as e:
                print(f"[!] Failed to fetch historic baselines: {e}")
                print("[?] Continuing structural pipeline execution, ignoring baseline fetching")

        if "sub_results" in meta_res:
            for sub_player, replaced_player in meta_res["sub_results"].items():
                s_low                       = sub_player.lower()
                chosen_team_id, chosen_tier = self.assignments[replaced_player.lower()]
                self.assignments[s_low]     = (chosen_team_id, chosen_tier)

                self.rosters[chosen_team_id]    .add(sub_player)
                self.subbed_players_set         .add(s_low)
                self.subbed_players_set         .add(replaced_player.lower())

                self.sub_relations[replaced_player.casefold()].append(sub_player)
                self.sub_relations[s_low] = [replaced_player]

        if not self.tour_label: self.tour_label = init_label

        self.val_str        = meta_res["th_str"]
        self.base_exp       = meta_res["base_exp"]
        self.new_players    = meta_res["selected_new"]
        self.dry_choice     = meta_res.get("dry_choice",    "No")
        self.share_choice   = meta_res.get("share_choice",  "No")
        self.exp_map        = {name: (self.base_exp - 1 if name.lower() in self.subbed_players_set else self.base_exp) for name in all_known}

        return True

    def get_stage_label(self) -> str:
        team_count = len(self.rosters) if self.use_teams else 0
        if team_count == 2: return "Final" if self.base_exp >= 3 else f"R{self.base_exp}"
        return "Mid-Tour" if self.base_exp == 3  else "Final" if (team_count <= 4 and self.base_exp >= 6) or (team_count > 4 and self.base_exp >= 5) else f"R{self.base_exp}"

    def process_and_generate(self):
        for path in self.json_paths:
            with open(path, encoding = "utf-8") as f: data = json.load(f)
            songs = data.get("songs", [])

            if not path.stem.startswith("amq"):
                match_digits = re.search(r"(\d+)$", path.stem)

                if match_digits:
                    m = int(match_digits.group(1))
                    if m <= THRESH_SONG: songs = songs[: min(m, len(songs))]

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

                if len(t_in_f) < 2 and len(self.rosters) >= 2:
                    all_team_ids    = list(self.rosters.keys())
                    missing_teams   = [f"Team {tid} ({self.t1_lookup.get(tid, tid)})" for tid in all_team_ids if tid not in t_in_f]

                    if missing_teams:
                        dialog = AskPlayerSelectionDialog(None, f"Missing Team in {path.name}", f"Only {len(t_in_f)} team(s) detected in {path.name}, which teams are actually playing?", missing_teams)
                        if dialog.result_selection:
                            sel_tid = int(dialog.result_selection.split()[1])
                            t_in_f.add(sel_tid)

                for tid in t_in_f:
                    for m_p in self.rosters[tid]:
                        if m_p.lower() not in self.assignments:
                            for c_p in raw_f_players:
                                if c_p.lower() in self.assignments and self.assignments[c_p.lower()][0] == tid: 
                                    self.assignments[m_p.lower()] = self.assignments[c_p.lower()]

                for tid in t_in_f:
                    team_roster     = self.rosters[tid]
                    present_in_team = {p for p in raw_f_players if p.lower() in self.assignments and self.assignments[p.lower()][0] == tid}

                    if len(present_in_team) < 4:
                        missing_candidates  = list(team_roster - present_in_team)
                        needed_count        = 4 - len(present_in_team)

                        if len(missing_candidates) > needed_count:
                            dialog = AskPlayerSelectionDialog(None, f"Ambiguous 0/0 Player in {path.name}", f"Which player is missing/playing for Team {self.t1_lookup.get(tid, tid)} in {path.name}?", sorted(missing_candidates))
                            if dialog.result_selection: final_members.add(dialog.result_selection)
                        else: final_members.update(missing_candidates)

            apply_rev       = len(final_members) % 2 == 0
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
                si          = song.get("songInfo",  {})
                st          = si.get("type",        3)
                t_num       = si.get("typeNumber",  0)
                ann_id      = str(si.get("annSongId"))
                is_chan     = ann_id in self.chanting_ids
                anime_name  = si.get("animeNames",  {}).get("romaji", "Unknown")
                song_name   = si.get("songName",    "Unknown")
                artist_name = si.get("artist",      "Unknown")

                if len(anime_name)  > THRESH_CHRL: anime_name   = re.sub(r"\s+\S*$", "", anime_name     [:THRESH_CHRL]) + " ..."
                if len(song_name)   > THRESH_CHRM: song_name    = re.sub(r"\s+\S*$", "", song_name      [:THRESH_CHRM]) + " ..."
                if len(artist_name) > THRESH_CHRS: artist_name  = re.sub(r"\s+\S*$", "", artist_name    [:THRESH_CHRS]) + " ..."

                type_fmt    = f"(OP{t_num})" if st == 1 else f"(ED{t_num})" if st == 2 else "(IN)"
                song_line   = f"{anime_name} {type_fmt}: {song_name} by {artist_name}"
                raw_correct = song.get("correctGuessPlayers", [])
                correct     = set()

                for p in raw_correct:
                    if      isinstance(p, str)                  : correct.add(p)
                    elif    isinstance(p, dict) and "name" in p : correct.add(p["name"])

                active_correct  = correct & final_members
                amt_correct     = len(active_correct)

                self.song_history.append((correct, raw_f_players))

                ls          = song.get("listStates", [])
                s_riggers   = {p["name"] for p in ls}

                self.global_stats["tot_c"] += len(correct)

                try:
                    vint_raw    = si.get("vintage", "")
                    yr          = int(extract_year(vint_raw)) if vint_raw else None
                    vint_scaled = float(extract_year(vint_raw)) if vint_raw else 0.0
                    vint_pretty = format_year(vint_scaled) if vint_raw else "Unknown"
                except Exception:
                    yr          = None
                    vint_pretty = "Unknown"

                if yr is not None: self.all_vint.append(yr)

                try                 : safe_diff = float(si.get("animeDifficulty", 0.0))
                except Exception    : safe_diff = 0.0

                self.all_diff   .append(safe_diff)
                self.song_data  .append({"vintage": yr if yr is not None else 0, "difficulty": safe_diff, "correct_count": int(len(correct))})

                song_line_hover = f"{anime_name} {type_fmt}: {song_name} by {artist_name} ({vint_pretty}/{safe_diff:.2f}: {len(correct)}/8)"

                if isinstance(si.get("animeGenre"), list): self.genre_c .update(si.get("animeGenre"))
                if isinstance(si.get("animeTags"),  list): self.tag_c   .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])

                if yr is not None and yr > 0:
                    diffs_arr   = [s["difficulty"] for s in self.song_data]
                    max_diff_v  = max(diffs_arr) if diffs_arr else 0
                    num_x_v     = 8 if max_diff_v < 40 else 9
                    num_y_v     = 8 if max_diff_v < 40 else 9
                    x_idx_v     = min(int(math.floor(safe_diff / 5)), num_x_v - 1)
                    vint_floor = math.floor(float(yr))

                    if num_y_v == 8 : y_idx_v = 0 if vint_floor < 1995 else min(int(math.floor((vint_floor - 1995) / 5)) + 1, 7)
                    else            : y_idx_v = 0 if vint_floor < 1990 else min(int(math.floor((vint_floor - 1990) / 5)) + 1, 8)

                    self.matrix_song_details[f"{x_idx_v}-{y_idx_v}"].append(song_line_hover)

                if len(correct) == 0: self.tour_song_details["Total 0/8s"].append(song_line)

                elif len(correct) == 1:
                    sw_v = list(correct)[0]
                    self.tour_song_details["Total 1/8s"].append(f"{song_line} ({sw_v})")
                    if sw_v.lower() in self.assignments: self.team_song_details[self.assignments[sw_v.lower()][0]]["Total 1/8s"].append(f"{song_line} ({sw_v})")

                elif len(correct) == 2:
                    p_list_v = list(correct)
                    self.tour_song_details["Total 2/8s"].append(f"{song_line} ({p_list_v[0]}/{p_list_v[1]})")

                elif apply_rev and len(final_members - correct) == 1:
                    missing_player_v = list(final_members - correct)[0]
                    self.tour_song_details["Total 7/8s"].append(f"{song_line} ({missing_player_v})")

                elif len(final_members - correct) == 0: self.tour_song_details["Total 8/8s"].append(song_line)

                if amt_correct == 0:
                    for name in final_members: self.p_zero_e[name] += 1

                if amt_correct == 1:
                    solo_winner = list(active_correct)[0]
                    if solo_winner not in s_riggers: self.p_offlist_erigs[solo_winner] += 1

                if self.use_teams and ls:
                    t_list = list({self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments})

                    if len(t_list) == 2:
                        tA, tB = t_list[0], t_list[1]

                        cA = correct & self.rosters[tA]
                        cB = correct & self.rosters[tB]

                        for r_item in ls:
                            r_name = r_item["name"]

                            if r_name.lower() in self.assignments:
                                r_team = self.assignments[r_name.lower()][0]

                                c_my_team   = cA if r_team == tA else cB
                                c_opp_team  = cB if r_team == tA else cA

                                if len(c_my_team) == 0 and len(c_opp_team) > 0: self.p_zero_x_rigs[r_name] += 1

                for sw_v in active_correct:
                    if amt_correct == 1: self.player_song_details[sw_v]["1/8s"].append(song_line)

                    elif amt_correct == 2:
                        opp_player_v    = (list(active_correct)[1] if sw_v.casefold() == list(active_correct)[0].casefold() and len(active_correct) > 1 else list(active_correct)[0])
                        t_sw_v          = self.assignments.get(sw_v         .lower(), (None,))[0] if self.use_teams else None
                        t_opp_v         = self.assignments.get(opp_player_v .lower(), (None,))[0] if self.use_teams else None

                        if t_sw_v is not None and t_opp_v is not None and t_sw_v == t_opp_v : self.player_song_details[sw_v]["2/8s"].append(f"{song_line} (covered by {opp_player_v})")
                        else                                                                : self.player_song_details[sw_v]["2/8s"].append(f"{song_line} (blocked by {opp_player_v})")

                    if safe_diff > 0    : self.p_hit_diff[sw_v].append(safe_diff)
                    if yr is not None   : self.p_hit_vint[sw_v].append(extract_year(si.get("vintage")))

                if amt_correct <= 3:
                    for sw_v in active_correct: self.p_three_or_below[sw_v] += 1

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
                            try                 : seen_song_times.add((str(p["name"]).casefold(), float(p["answerTime"])))
                            except Exception    : pass

                for key_name in ["answerTimes", "answerTime", "answerTimesByPlayer", "playerAnswerTimes"]:
                    val = song.get(key_name)

                    if isinstance(val, dict):
                        for p_name, t_val in val.items():
                            try                             : seen_song_times.add((str(p_name).casefold(), float(t_val)))
                            except (ValueError, TypeError)  : pass

                name_map = {m.lower(): m for m in final_members}

                for p_name_lower, t_float in seen_song_times:
                    if p_name_lower in name_map: self.p_answer_times[name_map[p_name_lower]].append(t_float)

                if len(ls) == 1:
                    u                   =   ls[0]["name"]
                    self.p_l_solos[u]   +=  1

                    if not (len(correct) == 1 and list(correct)[0] == u): self.p_m_erigs[u] += 1

                if ls:
                    is_true_solo_rig = len(ls) == 1

                    for p in ls:
                        n_v         = p["name"]
                        marker_v    = "✓" if (n_v in active_correct) else "✗"

                        self.player_song_details[n_v]["Rigs"].append(f"{marker_v} {song_line}")

                        if is_true_solo_rig:
                            self.player_song_details[n_v]["Solo Rigs"].append(f"{marker_v} {song_line}")

                            s_v = sorted(list(active_correct - {n_v}))
                            sC_v = len(s_v)
                            tag_v = (
                                "(0/8)"
                                if sC_v == 0 else f"(stolen by {s_v[0]})"
                                if sC_v == 1 else f"(stolen by {s_v[0]}/{s_v[1]})"
                                if sC_v == 2 else f"({amt_correct}/8)"
                            )

                            if n_v in active_correct and amt_correct == 1   : self.player_song_details[n_v]["Solo Rig Conversions"].append(f"✓ {song_line}")
                            else                                            : self.player_song_details[n_v]["Solo Rig Conversions"].append(f"✗ {song_line} {tag_v}")

                if self.use_teams:
                    t_list = list({self.assignments[p.lower()][0] for p in raw_f_players if p.lower() in self.assignments})

                    if len(t_list) == 2:
                        tA, tB = t_list[0], t_list[1]
                        cA, cB = correct & self.rosters[tA], correct & self.rosters[tB]

                        if len(cA) == 4 and not cB:
                            self.t_sweeps[tA]           += 1
                            self.global_stats["sweeps"] += 1

                        if len(cB) == 4 and not cA:
                            self.t_sweeps[tB]           += 1
                            self.global_stats["sweeps"] += 1

                        if (len(cA & final_members) == 4 and not (cB & final_members)) or (
                            len(cB & final_members) == 4 and not (cA & final_members)): 
                            self.tour_song_details["Total 4-0s"].append(song_line)

                        for cur, opp in [(tA, tB), (tB, tA)]:
                            cC, oC = correct & self.rosters[cur], correct & self.rosters[opp]

                            if not oC:
                                for p in cC: self.p_pts[p] += 1

                            if len(cC) == 1 and len(oC) > 0:
                                lone_p = list(cC)[0]
                                self.p_blks[lone_p] += 1

                        for _, opp_v, cC_v, oC_v in [(tA, tB, cA & final_members, cB & final_members), (tB, tA, cB & final_members, cA & final_members)]:
                            oL_v = self.t1_lookup.get(opp_v, f"Team {opp_v}")

                            if not oC_v:
                                for p_v in cC_v: self.player_song_details[p_v]["Lives Taken"].append(f"{song_line} (from Team {oL_v})")

                            if len(cC_v) == 1 and len(oC_v) > 0:
                                oP_v = sorted(list(oC_v), key=lambda x: self.assignments.get(x.lower(), (None, "5"))[1])

                                opp_tag_v = (
                                    f"(from {oP_v[0]} in Team {oL_v})"
                                    if len(oP_v) == 1 and oP_v[0] != oL_v                       else f"(from {oP_v[0]})"
                                    if len(oP_v) == 1                                           else f"(from {oP_v[0]}/{oP_v[1]} in Team {oL_v})"
                                    if len(oP_v) == 2 and oP_v[0] != oL_v and oP_v[1] != oL_v   else f"(from {oP_v[0]}/{oP_v[1]})"
                                    if len(oP_v) == 2                                           else f"(from Team {oL_v})"
                                )

                                self.player_song_details[list(cC_v)[0]]["Lives Saved"].append(f"{song_line} {opp_tag_v}")

                    for tid in t_list:
                        ros     = self.rosters[tid]
                        c_on_t  = correct & ros

                        self.t_c_ps[tid].append(len(c_on_t) / 4.0)
                        if yr is not None: self.t_vint[tid].append(yr)

                        if s_riggers & ros:
                            self.t_on_syn       [tid].append(len(c_on_t)                / 4.0)
                            self.t_sh_rig       [tid].append((len(s_riggers & ros) - 1) / 3.0)
                        else: self.t_off_syn    [tid].append(len(c_on_t)                / 4.0)

                if len(final_members - correct) == 0: self.global_stats["fulls"] += 1

                elif apply_rev and len(final_members - correct) == 1:
                    self.global_stats["sevens"]                     += 1
                    self.p_rev_e[list(final_members - correct)[0]]  += 1

                elif len(correct) == 2:
                    self.global_stats["doubles"] += 1

                    p_list = list(correct)
                    p1, p2 = p_list[0], p_list[1]

                    self.p_two_e[p1] += 1
                    self.p_two_e[p2] += 1

                elif len(correct) == 1:
                    self.global_stats["solos"]  +=  1
                    sw                          =   list(correct)[0]
                    self.e_counts[sw]           +=  1
                    if sw.lower() in self.assignments: self.t_solos[self.assignments[sw.lower()][0]] += 1

                elif len(correct) == 0: self.global_stats["blanks"] += 1

                amtcorrect = len(correct)

                if amtcorrect > 0:
                    teamsize    = 4
                    uf_song     = sum(math.comb(2 * teamsize - (i + 2), amtcorrect - 1) / math.comb(2 * teamsize - 1, amtcorrect - 1) for i in range(teamsize)) / teamsize

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

                        if yr is not None: self.p_c_vint[name]      .append(yr)
                        self.player_song_details[name]["Overall"]   .append(f"✓ {song_line}")

                    else:
                        if st in [1, 2, 3]  : self.player_song_details[name][f"Type {st}"]  .append(f"✗ {song_line}")
                        if is_chan          : self.player_song_details[name]["Chant"]       .append(f"✗ {song_line}")

                        self.player_song_details[name]["Overall"].append(f"✗ {song_line}")

                    if is_chan: self.p_chan_s[name] += 1

                if ls:
                    for p in ls:
                        n               =   p["name"]
                        self.p_rigs[n]  +=  1

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

        stage   = self.get_stage_label()
        prefix  = f"{self.tour_label.strip()} Tour: "

        png_path = self.tour_dir / "png"
        web_path = self.tour_dir / "site"

        for path_dir in [png_path, web_path]:
            path_dir.mkdir(parents = True, exist_ok = True)
            for item in path_dir.iterdir():
                if item.is_file(): item.unlink()

        watched_valid = self.missing_list_count <= THRESH_WTCH

        tasks = [
            (create_player_png,     (self, self.elo_map, watched_valid, stage, png_path, self.apps, prefix, self.exp_map, self.base_exp, self.new_players, self.val_str)),
            (create_tour_png,       (self, self.use_teams, watched_valid, png_path)),
            (create_scatter_png,    (self, png_path, False, self.elo_map)),
            (create_song_png,       (self, png_path)),
            (create_dashboard_html, (self, web_path, self.use_teams, watched_valid)),
        ]

        if self.assignments:
            tasks.append((create_tier_png, (self, self.assignments, png_path, any(self.p_chan_s.values()))))
            if watched_valid: tasks.append((create_team_png, (self, self.assignments, self.t1_lookup, png_path)))

        if watched_valid: tasks.append((create_scatter_png, (self, png_path, True, self.elo_map)))

        with concurrent.futures.ProcessPoolExecutor() as executor:
            task_map = {executor.submit(func, *args): func.__name__ for func, args in tasks}

            for future in concurrent.futures.as_completed(task_map):
                task_name = task_map[future]

                try                     : future.result()
                except Exception as e   : print(f"Task {task_name} failed: {e}")

        fuse_images(png_path)
        workspace_root = self.script_dir.parent

        target_hako_dir = workspace_root / f"hako_{self.tour_id}"
        target_hako_dir.mkdir(parents = True, exist_ok = True)

        allowed_files   = {"Player.png", "Extra.png"}
        player_file     = png_path / "Player.png"
        extra_file      = png_path / "Extra.png"

        if player_file.exists():
            shutil.copy(player_file, target_hako_dir / "Player.png")
            print(f"[✓] Copied Player.png to {target_hako_dir.name}/Player.png")

        if extra_file.exists():
            shutil.copy(extra_file, target_hako_dir / "Extra.png")
            print(f"[✓] Copied Extra.png to {target_hako_dir.name}/Extra.png")

        if web_path.exists():
            zip_dest_path = target_hako_dir / "Site.zip"

            with zipfile.ZipFile(zip_dest_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root_dir, _, files in os.walk(web_path):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname   = os.path.relpath(file_path, web_path)

                        zipf.write(file_path, arcname)

            print(f"[✓] Zipped site contents to {target_hako_dir.name}/Site.zip")

        for file_path in png_path.glob("*.png"):
            if file_path.name not in allowed_files:
                try                 : file_path.unlink()
                except Exception    : pass

        if      self.dry_choice     == "Yes"    and getattr(self, "mode_choice", "Tour") == "Ant"   : self._handle_ant_spreadsheet_push ()
        elif    self.dry_choice     != "No"                                                         : self._handle_dry_script_execution ()
        if      self.share_choice   == "Yes"                                                        : self._handle_github_deploy        (web_path)

    def _handle_dry_script_execution(self):
        target_jsons_dir    = self.script_dir.parent / "jsons"
        target_codes_file   = self.script_dir.parent / "codes.txt"

        source_jsons_dir    = self.tour_dir / "json"
        source_codes_file   = self.tour_dir / "code.txt"

        script_name     = "ngm_local.py" if "ngm_local" in self.dry_choice else "ngm_stats.py"
        target_script   = self.script_dir.parent / script_name

        print(f"[?] Processing Tour {self.tour_id} using Dry's script")

        for file in self.script_dir     .glob("*.png")  : file.unlink()
        for file in target_jsons_dir    .glob("*.json") : file.unlink()

        if source_jsons_dir.exists():
            for file in source_jsons_dir.glob("*.json"): shutil.copy(file, target_jsons_dir / file.name)

        print("[✓] Copied JSONs to Dry's workspace")

        if source_codes_file.exists():
            shutil.copy(source_codes_file, target_codes_file)
            print("[✓] Copied code.txt to Dry's workspace")

        print("[?] Running Dry's script")

        try:
            subprocess.run([sys.executable, str(target_script)], cwd=str(self.script_dir.parent), check=True)
            print("[✓] Ran Dry's script successfully")

            output_dir = self.tour_dir / "dry"
            output_dir.mkdir(parents = True, exist_ok = True)

            files_to_copy = {
                "Stats.png"                         : "1-Player.png",
                "Stats2.png"                        : "2-Type.png",
                "Stats3 - Watched Exclusive.png"    : "3-List.png",
                "Stats Songs.png"                   : "4-Song.png",
                "Stats4.png"                        : "5-Extra.png",
            }

            print("[?] Copying Dry's PNGs back")

            for src_name, dest_name in files_to_copy.items():
                src_file    = self.script_dir.parent    / src_name
                dest_file   = output_dir                / dest_name

                if src_file.exists():
                    shutil.copy(src_file, dest_file)
                    print(f"[✓] Copied {src_name} as {dest_name}")
                else: print(f"[X] {src_name} not found in Dry's workspace")

        except subprocess.CalledProcessError as e: print(f"[X] Failed to run Dry's script: {e}")

    def _handle_ant_spreadsheet_push(self):
        print("[?] Pushing Ant stats to Google Spreadsheet")
        ANT_SHEET_ID = "1R1Th9ngAr5RwQxX5KforK8xDzplRXGbcF1axCE8liF4"

        if      self.tour_label == "Usual"                  : gid = 0
        elif    self.tour_label == "Watched"                : gid = 220235184
        elif    self.tour_label == "Watched OP"             : gid = 313246506
        elif    self.tour_label == "Watched IN"             : gid = 1251660941
        elif    self.tour_label == "Watched IN -Chanting"   : gid = 559526622
        elif    self.tour_label == "Watched 2+8s"           : gid = 881563885
        elif    self.tour_label == "Watched 2010+"          : gid = 548559487
        elif    self.tour_label == "Random OP"              : gid = 1955531566
        elif    self.tour_label == "Random IN"              : gid = 71038376
        else                                                : gid = 1085890115

        cred_file = self.script_dir / DIR_CREDS / "ant_credentials.json"
        auth_file = self.script_dir / DIR_CREDS / "ant_authorized_user.json"

        try:
            gc              = gspread.oauth(credentials_filename = str(cred_file), authorized_user_filename = str(auth_file))
            sheet           = gc.open_by_key(ANT_SHEET_ID)
            wks             = sheet.get_worksheet_by_id(gid)
            iso_timestamp   = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            valid_elos      = [float(v) for v in self.elo_map.values() if str(v).replace(".", "", 1).isdigit() or (str(v).startswith("-") and str(v)[1:].replace(".", "", 1).isdigit())]
            avg_rank        = np.mean(valid_elos) if valid_elos else 1.0
            df_teams        = compute_team_rows(self, self.assignments, self.t1_lookup)
            team_wlt_map    = {}

            if not df_teams.empty:
                for _, t_row in df_teams.iterrows():
                    tid       = t_row["_tid"]
                    h_payload = t_row["_history"]

                    if h_payload["total_matches"] > 0:
                        w_cnt = sum(len(v) for v in h_payload["wins"]   .values())
                        l_cnt = sum(len(v) for v in h_payload["losses"] .values())
                        t_cnt = sum(len(v) for v in h_payload["ties"]   .values())

                        team_wlt_map[tid] = (w_cnt, l_cnt, t_cnt)

            def player_sort_key(x):
                gr = (self.c_counts[x] / self.s_part[x]) if self.s_part[x] else 0.0

                try                 : elo = float(self.elo_map.get(x.lower(), float("inf")))
                except Exception    : elo = float("inf")

                return (gr, -elo)

            sorted_players  = sorted(self.s_part.keys(), key = player_sort_key, reverse = True)
            rows_to_push    = [[""] * 32]

            for i, name in enumerate(sorted_players):
                tot = self.s_part[name]
                cor = self.c_counts[name]
                gr  = (cor / tot * 100) if tot else 0.0

                uf_val      = (self.p_usefulness_sum    [name] * avg_rank * 8) / tot    if tot else 0.0
                erigs       = self.e_counts             [name]
                got_78      = self.p_rev_e              [name]
                avg_8       = (self.p_overs_sum         [name] / cor)                   if cor else ""
                three_below = self.p_three_or_below     [name]

                op_gr   = (self.p_type_c[name][1] / self.p_type_s[name][1] * 100) if self.p_type_s[name][1] else ""
                ed_gr   = (self.p_type_c[name][2] / self.p_type_s[name][2] * 100) if self.p_type_s[name][2] else ""
                in_gr   = (self.p_type_c[name][3] / self.p_type_s[name][3] * 100) if self.p_type_s[name][3] else ""

                avg_diff = np.mean      (self.p_hit_diff        [name]) if self.p_hit_diff      [name] else ""
                med_time = np.median    (self.p_answer_times    [name]) if self.p_answer_times  [name] else ""

                rigs            = self.p_rigs           [name]
                rigs_h          = self.p_rigs_h         [name]
                onlist          = (rigs_h / rigs * 100)                 if rigs                 else ""
                offlist         = ((cor - rigs_h) / (tot - rigs) * 100) if (tot - rigs)         else ""
                rig_pct         = (rigs / tot * 100)                    if tot                  else ""
                solo_rigs       = self.p_l_solos        [name]
                rigs_missed     = rigs - rigs_h                         if rigs                 else ""
                avg_8_rigs      = np.mean(self.p_l_corr [name])         if self.p_l_corr[name]  else ""
                zero_e          = self.p_zero_e         [name]
                lives_taken     = self.p_pts            [name]          if self.use_teams       else ""
                lives_saved     = self.p_blks           [name]          if self.use_teams       else ""
                missed_solos    = self.p_m_erigs        [name]          
                zero_x_rigs     = self.p_zero_x_rigs    [name]          if self.use_teams       else ""
                offlist_erigs   = self.p_offlist_erigs  [name]

                win_val, lose_val, tie_val = "", "", ""

                if self.use_teams and name.lower() in self.assignments:
                    tid = self.assignments[name.lower()][0]

                    if tid in team_wlt_map:
                        w_c, l_c, t_c = team_wlt_map[tid]

                        win_val  = w_c
                        lose_val = l_c
                        tie_val  = t_c

                row = [
                    iso_timestamp           if i            == 0    else "",    # Timestamp
                    name,                                                       # Name
                    round(gr,       2)      if tot                  else "",    # GR
                    round(uf_val,   2)      if uf_val               else "",    # UF
                    erigs,                                                      # 1/8
                    zero_e,                                                     # 0/8
                    got_78,                                                     # 7/8
                    round(avg_8,        2)  if avg_8        != ""   else "",    # Mean Over-8
                    three_below,                                                # <=3/8
                    round(op_gr,        2)  if op_gr        != ""   else "",    # OPGR
                    round(ed_gr,        2)  if ed_gr        != ""   else "",    # EDGR
                    round(in_gr,        2)  if in_gr        != ""   else "",    # INGR
                    lives_taken,                                                # Lives Taken
                    lives_saved,                                                # Lives Saved
                    cor,                                                        # Corrects
                    tot,                                                        # Total
                    round(avg_diff,     2)  if avg_diff     != ""   else "",    # Mean Difficulty Hit
                    round(med_time,     2)  if med_time     != ""   else "",    # Median Time
                    win_val,                                                    # Wins
                    lose_val,                                                   # Losses
                    tie_val,                                                    # Ties
                    round(onlist,       2)  if onlist       != ""   else "",    # Rig GR
                    round(offlist,      2)  if offlist      != ""   else "",    # Off GR
                    round(rig_pct,      2)  if rig_pct      != ""   else "",    # Rig Rate
                    rigs,                                                       # Rigs
                    solo_rigs,                                                  # Solo Rigs
                    missed_solos,                                               # Missed Solos
                    rigs_h                  if rigs                 else "",    # Rigs Hit
                    rigs_missed             if rigs                 else "",    # Rigs Missed
                    zero_x_rigs,                                                # Lives Lost on Rigs
                    offlist_erigs,                                              # Off 1/8
                    round(avg_8_rigs,   2)  if avg_8_rigs   != ""   else ""     # Rig Over-8
                ]

                rows_to_push.append(row)

            existing_values = wks.get_all_values()
            last_data_row   = len(existing_values)

            while last_data_row > 0 and not any(existing_values[last_data_row - 1]): last_data_row -= 1
            insert_start_row = last_data_row + 1
            wks.insert_rows(rows_to_push, row = insert_start_row, value_input_option = "USER_ENTERED")
            print("[✓] Successfully pushed Ant stats to Google Spreadsheet!")
        except Exception as e: print(f"[X] Failed to push Ant stats to Google Spreadsheet: {e}")

    def _handle_github_deploy(self, web_path: Path):
        timestamp   = datetime.datetime.now().strftime("%y%m%d%H%M")
        archive_dir = self.script_dir.parent / "hako" / "archive" / timestamp

        print(f"[?] Copying files to archive/{timestamp}")

        try:
            shutil.copytree(web_path, archive_dir, dirs_exist_ok=True)
            print(f"[✓] Copied all files from {web_path.name}")
        except Exception as e: print(f"[X] Failed to copy files: {e}")

        dashboard_url = f"https://frittutisna.github.io/Stats-Maker/hako/archive/{timestamp}/index.html?update=1"
        print("[?] Pushing to GitHub")

        try:
            subprocess.run(["git", "add",       "."],                   check = True)
            subprocess.run(["git", "commit",    "-m", "Updated tour"],  check = True)
            subprocess.run(["git", "push"],                             check = True)

            print(f"[✓] Deployment completed, dashboard link: {dashboard_url}")
        except subprocess.CalledProcessError as git_error: print(f"[X] Failed to push to GitHub: {git_error}")