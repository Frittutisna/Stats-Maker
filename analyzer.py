import hashlib, json, logging, math, matplotlib, os, re, shutil, sys, time

logging     .getLogger  ("adjustText").setLevel(logging.ERROR)
matplotlib  .use        ('Agg')

import concurrent.futures       as fut
import matplotlib.colors        as mc
import matplotlib.pyplot        as plt
import numpy                    as np
import pandas                   as pd

from adjustText     import adjust_text
from collections    import Counter, defaultdict
from help.config    import *
from help.dialog    import *
from html2image     import Html2Image
from pathlib        import Path
from PIL            import Image
from scipy.spatial  import ConvexHull
from tkinter        import messagebox

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
        self.p_m_erigs                  = defaultdict(int)
        self.p_l_solos                  = defaultdict(int)
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
        if not loaded: return False

        self.use_teams, self.elo_map, self.assignments, self.t1_lookup, self.rosters, all_known, sub_candidates_raw, original_players_display = loaded

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
            if      init_label == "Watched 2+8"                 : default_th = "25, 20, 15, 10, 5"
            elif    init_label in ["Watched",   "QuagWatched"]  : default_th = "28, 18, 12, 6"
            elif    init_label in ["Usual",     "Quagsual"]     : default_th = "28, 19, 8"
            else                                                : default_th = "28, 19, 8"

        meta_dialog = TourMetadataDialog(None, self.tour_id, init_label, default_th, baseline_initial, list(all_known), self.elo_map, sub_candidates_raw, original_players_display, self.tour_dir)
        if meta_dialog.result is None: sys.exit(0)

        meta_res        = meta_dialog.result
        self.tour_label = meta_res["tour_label"]

        if "sub_results" in meta_res:
            for sub_player, replaced_player in meta_res["sub_results"].items():
                s_low = sub_player.lower()
                chosen_team_id, chosen_tier = self.assignments[replaced_player.lower()]
                self.assignments[s_low] = (chosen_team_id, chosen_tier)
                self.rosters[chosen_team_id].add(sub_player)
                self.subbed_players_set.add(s_low)
                self.subbed_players_set.add(replaced_player.lower())
                self.sub_relations[replaced_player.casefold()].append(sub_player)
                self.sub_relations[s_low] = [replaced_player]

        if not self.tour_label: self.tour_label = init_label

        self.val_str        = meta_res["th_str"]
        self.base_exp       = meta_res["base_exp"]
        self.new_players    = meta_res["selected_new"]
        self.exp_map        = {}
        mismatched_players  = {}

        for name in list(all_known):
            act = len(self.apps.get(name, []))

            if act < self.base_exp  : mismatched_players[name]  = act
            else                    : self.exp_map[name]        = self.base_exp

        if mismatched_players:
            mismatch_dialog = MismatchedRoundsDialog(None, mismatched_players, self.base_exp, self.subbed_players_set, self.tour_dir, watched_valid)
            if mismatch_dialog.result is None: sys.exit(0)
            mismatch_res = mismatch_dialog.result

            for name, target in mismatch_res.items():
                act                 = len(self.apps.get(name, []))
                self.exp_map[name]  = target

                if target == "ignore": continue

                if target > act:
                    avg_songs_per_json  = sum(len(self.apps.get(n, [])) for n in all_known) / len(all_known)
                    missing_rounds      = target - act
                    self.s_part[name]   += int(missing_rounds * avg_songs_per_json)

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

                song_line = f"{anime_name} {type_fmt}: {song_name} by {artist_name}"

                if isinstance(si.get("animeGenre"), list): self.genre_c .update(si.get("animeGenre"))
                if isinstance(si.get("animeTags"),  list): self.tag_c   .update([t for t in si.get("animeTags") if t not in EXCLUDED_TAGS])

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
                    vint_raw    = si.get("vintage", "")
                    yr          = int(extract_year(vint_raw)) if vint_raw else None

                except: yr = None

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

                if yr is not None and yr > 0:
                    diffs_arr   = [s["difficulty"] for s in self.song_data]
                    max_diff_v  = max(diffs_arr) if diffs_arr else 0
                    num_x_v     = 8 if max_diff_v < 40 else 9
                    num_y_v     = 8 if max_diff_v < 40 else 9
                    x_idx_v     = min(int(math.floor(safe_diff / 5)), num_x_v - 1)
                    vint_floor  = math.floor(float(yr))

                    if num_y_v == 8 : y_idx_v = 0 if vint_floor < 1990 else min(int(math.floor((vint_floor - 1990) / 5)) + 1, 7)
                    else            : y_idx_v = min(max(int(math.floor((vint_floor - 1985) / 5)), 0), 8)

                    self.matrix_song_details[f"{x_idx_v}-{y_idx_v}"].append(song_line)

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
                        opp_player_v = list(active_correct)[1] if sw_v.casefold() == list(active_correct)[0].casefold() and len(active_correct) > 1 else list(active_correct)[0]
                        t_sw_v       = self.assignments.get(sw_v.lower(), (None,))[0] if self.use_teams else None
                        t_opp_v      = self.assignments.get(opp_player_v.lower(), (None,))[0] if self.use_teams else None

                        if t_sw_v is not None and t_opp_v is not None and t_sw_v == t_opp_v : self.player_song_details[sw_v]["2/8s"].append(f"{song_line} (covered by {opp_player_v})")
                        else                                                                : self.player_song_details[sw_v]["2/8s"].append(f"{song_line} (blocked by {opp_player_v})")

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
                                for p in cC: self.p_pts[p] += 1

                            if len(cC) == 1 and len(oC) > 0: self.p_blks[list(cC)[0]] += 1

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

                        if n in correct     : self.p_rigs_h [n] += 1
                        if yr is not None   : self.p_l_vint [n].append(yr)

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
        out_path    = self.tour_dir / DIR_OUT

        out_path.mkdir(parents = True, exist_ok = True)
        for item in out_path.iterdir(): item.unlink()

        tasks = []

        tasks.append((self._create_player_png,      (self.elo_map, watched_valid, stage, out_path, self.apps, prefix, self.exp_map, self.base_exp, self.new_players, self.val_str)))
        tasks.append((self._create_tour_png,        (self.use_teams, watched_valid, out_path)))
        tasks.append((self._create_scatter_png,     (out_path, False, self.elo_map)))
        tasks.append((self._create_song_png,        (out_path, )))
        tasks.append((self._create_dashboard_html,  (out_path, self.use_teams, watched_valid)))

        if self.assignments:
            tasks.append((self._create_tier_png, (self.assignments, out_path, any(self.p_chan_s.values()))))
            if watched_valid: tasks.append((self._create_team_png, (self.assignments, self.t1_lookup, out_path)))

        if watched_valid: tasks.append((self._create_scatter_png, (out_path, True, self.elo_map)))

        with fut.ProcessPoolExecutor() as executor:
            task = {executor.submit(func, *args): func.__name__ for func, args in tasks}

            for future in fut.as_completed(task):
                task_name = task[future]

                try                     : future.result()
                except Exception as e   : print(f"Task {task_name} failed: {e}")

        self._fuse(out_path)
        allowed_files = {"General.png", "Player.png", "Extra.png", "Plots.png"}

        for file_path in out_path.glob("*.png"):
            if file_path.name not in allowed_files:
                try                 : file_path.unlink()
                except Exception    : pass

        print(f"Push to GitHub to update the online Dashboard: https://frittutisna.github.io/Stats-Maker/tour/{self.tour_id}/hakohoka/Dashboard.html?update=1")

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
        if not codes.exists() or os.path.getsize(codes) == 0: return False, {}, {}, {}, defaultdict(set), all_known
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
        alias_path                                  = self.script_dir / DIR_TOURS / FILE_ALIAS
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

        unresolved_players = [p for p in all_known if p.lower() not in assignments]

        if unresolved_players:
            messagebox.showerror("Roster Mismatch", f"These players are in the JSONs but not in codes.txt: {', '.join(unresolved_players)}")
            return False

        original_players_display = [p for p in all_known if p.lower() in assignments and p.lower() not in [s.lower() for s in sub_candidates_raw]]

        if new_aliases:
            existing_entries = []

            if alias_path.exists():
                with open(alias_path, "r", encoding = "utf-8") as f: existing_entries = [l.strip().lower() for l in f.readlines()]

            with open(alias_path, "a", encoding = "utf-8") as f:
                for k, v in new_aliases.items():
                    entry = f"{k}, {v}".lower()

                    if entry not in existing_entries:
                        f.write(f"{k}, {v}\n")
                        existing_entries.append(entry)

        return True, elo_map, assignments, t1_lookup, rosters, all_known, sub_candidates_raw, original_players_display

    def _compute_player_rows(self, elo_map, apps, exp_map, base_exp, new_players, watched, active, t_labels, avg_rank):
        rows, eligibility = [], []

        for name in self.s_part:
            tot, cor    = self.s_part[name], self.c_counts[name]
            target      = exp_map.get(name, base_exp)
            d_name      = name

            if name in new_players: d_name += " ☆"

            if target != "ignore" and target < base_exp:
                if name.lower() in self.main_roster_names   : d_name += " ▼"
                else                                        : d_name += " ▲"

            is_eligible = not ("▼" in d_name or "▲" in d_name)
            eligibility.append(is_eligible)
            act = len(apps.get(name, []))

            if target != "ignore" and act < target:
                syms = ["", "(1)", "(2)", "(3)", "(4)", "(5)", "(6)"]
                if 0 < (target-act) < len(syms): d_name += f" {syms[target-act]}"

            row = {"Player": d_name}
            if self.use_teams: row["Elo"] = elo_map.get(name.lower(), np.nan)
            row.update({"Guess Rate": cor / tot if tot else 0.0})

            if self.use_teams: 
                uf_val = (self.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0
                row.update({"UF": uf_val})
                
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
                row[t_labels[tid]]  = self.p_type_c[name][tid] / seen if seen else np.nan

            if watched:
                rig_over8 = np.mean(self.p_l_corr[name]) if self.p_l_corr[name] else np.nan

                row.update({
                    "Rigs"              : self.p_rigs[name],
                    "Rig Rate"          : self.p_rigs[name]             / tot                       if tot                          else np.nan,
                    "Solo Rigs"         : self.p_l_solos[name],
                    "Solo Rig Rate"     : self.p_l_solos[name]          / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Rig Over-8"        : rig_over8,
                    "Over-8 Delta"      : rig_over8 - avg_over8,
                    "Rig Guess Rate"    : self.p_rigs_h[name]           / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Off Guess Rate"    : (cor - self.p_rigs_h[name])   / (tot - self.p_rigs[name]) if (tot - self.p_rigs[name])    else np.nan,
                    "Rig Delta"         : (cor - self.p_rigs[name])     / cor                       if cor                          else np.nan,
                })

            times       = self.p_answer_times.get(name, [])
            seen_chan   = self.p_chan_s[name]

            row["Median Time"]      = np.median(times) if times else np.nan
            row["Chant Guess Rate"] = self.p_chan_c[name] / seen_chan if seen_chan else np.nan

            rows.append(row)

        df = pd.DataFrame(rows)

        if "Score" in df.columns: df = df.sort_values(by = ["Guess Rate", "Score"], ascending = [False, False])

        elif "Elo" in df.columns:
            df["_sort_elo"] = pd.to_numeric(df["Elo"], errors = 'coerce')
            df              = df.sort_values(by = ["Guess Rate", "_sort_elo"], ascending = [False, True]).drop(columns = ["_sort_elo"])

        else: df = df.sort_values("Guess Rate", ascending = False)

        mask = pd.Series(eligibility, index = pd.DataFrame(rows).index).reindex(df.index).values
        return df, mask

    def _create_player_png(self, elo_map, watched, stage, path, apps, prefix, exp_map, base_exp, new_players, val_str):
        t_labels    = {1: "OP Guess Rate", 2: "ED Guess Rate", 3: "IN Guess Rate"}
        active      = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]

        if len(active) <= 1 : active = []

        valid_elos  = [float(v) for v in elo_map.values() if str(v).replace('.', '', 1).isdigit() or (str(v).startswith('-') and str(v)[1:].replace('.', '', 1).isdigit())]
        avg_rank    = np.mean(valid_elos) if valid_elos else 1.0
        df, mask    = self._compute_player_rows(elo_map, apps, exp_map, base_exp, new_players, watched, active, t_labels, avg_rank)
        df_png      = df.copy()
        pcts        = ["Guess Rate"] + [t_labels[t] for t in active] + (["Rig Rate", "Solo Rig Rate", "Rig Delta", "Rig Guess Rate", "Off Guess Rate"] if watched else []) + ["Chant Guess Rate"]

        if "Elo"            in df_png.columns: df_png["Elo"]            = pd.to_numeric(df_png["Elo"],          errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "UF"             in df_png.columns: df_png["UF"]             = pd.to_numeric(df_png["UF"],           errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Score"          in df_png.columns: df_png["Score"]          = pd.to_numeric(df_png["Score"],        errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Median Time"    in df_png.columns: df_png["Median Time"]    = pd.to_numeric(df_png["Median Time"],  errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Mean Over-8"    in df_png.columns: df_png["Mean Over-8"]    = pd.to_numeric(df_png["Mean Over-8"],  errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Rig Over-8"     in df_png.columns: df_png["Rig Over-8"]     = pd.to_numeric(df_png["Rig Over-8"],   errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Over-8 Delta"   in df_png.columns: df_png["Over-8 Delta"]   = pd.to_numeric(df_png["Over-8 Delta"], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

        for c in pcts: df_png[c] = pd.to_numeric(df_png[c], errors = 'coerce').mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
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
            ["Mean Guess Rate", f"{100 * (self.global_stats['tot_c'] / sum(self.s_part.values())):.2f}" if self.s_part      else "N/A", None],
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

    def _compute_team_rows(self, assigns, t1_lookup):
        res = []

        for tid in self.t_c_ps:
            t_overs = []

            for original_name in self.s_part:
                n_lower = original_name.lower()

                if n_lower in assigns:
                    t_info = assigns[n_lower]
                    if t_info[0] == tid and self.c_counts[original_name] > 0: t_overs.append(self.p_overs_sum[original_name] / self.c_counts[original_name])
         
            t_elos = []

            for p in self.rosters[tid]:
                v = self.elo_map.get(p.lower())

                if v is not None:
                    try     : t_elos.append(float(v))
                    except  : pass

            res.append({
                "Team Leader"   : t1_lookup.get(tid, f"Team {tid}"),
                "Mean Elo"      : np.mean(t_elos),
                "Mean GR"       : np.mean(self.t_c_ps       [tid]) * 100,
                "Total 1/8s"    : self.t_solos              [tid],
                "Mean Over-8"   : np.mean(t_overs),
                "Rig Synergy"   : np.mean(self.t_on_syn     [tid]) * 100,
                "Off Synergy"   : np.mean(self.t_off_syn    [tid]) * 100,
                "Shared Rigs"   : np.mean(self.t_sh_rig     [tid]) * 100,
                "_tid"          : tid
            })

        df = pd.DataFrame(res).sort_values(by = ["Mean GR", "Mean Elo"], ascending = [False, True])
        return df

    def _create_team_png(self, assigns, t1_lookup, path):
        df          = self._compute_team_rows(assigns, t1_lookup)
        df_png      = df.drop(columns = ["_tid"])
        num_cols    = ["Mean Elo", "Mean GR", "Mean Over-8", "Rig Synergy", "Off Synergy", "Shared Rigs"]

        for c in num_cols: df_png[c] = pd.to_numeric(df_png[c], errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        self._export_png(df_png, path, "Team.png", "Team Statistics")

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

            row1["Guess Rate"]          = f"{gen_players[0]['player']} ({gen_players[0]['value']:.2f})" if gen_players                                  else "N/A"
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
                    f"<td{s_gen}>{row['Guess Rate']}</td>"
                    f"<td{s_atk}>{row['Lives Taken']}</td>"
                    f"<td{s_blk}>{row['Lives Saved']}</td>"
                f"</tr>"
            )

        html_parts.append(
            "<tr>"
                "<th style='border-top: 3px solid black;'>Tier</th>"
                "<th style='border-top: 3px solid black;'>Contribution Rate</th>"
                "<th style='border-top: 3px solid black;'>Median Time</th>"
                "<th style='border-top: 3px solid black;'>Chanting Guess Rate</th>"
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

            if name in self.new_players  : d_name += " ☆"

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

            act = len(self.apps.get(name, []))

            if target != "ignore" and act < target:
                syms = ["", "(1)", "(2)", "(3)", "(4)", "(5)", "(6)"]
                if 0 < (target-act) < len(syms): d_name += f" {syms[target-act]}"

            row = {"Player": {"count": d_name, "details": [sub_hover] if sub_hover else []}}

            if self.use_teams: 
                try     : row["Elo"] = float(self.elo_map.get(name.lower(), np.nan))
                except  : row["Elo"] = np.nan

            if row_data is not None:
                tot, cor = self.s_part[name], self.c_counts[name]
                for key_details in ["Overall", "Type 1", "Type 2", "Type 3", "Chant"]: self.player_song_details[name][key_details].sort(key = lambda s: s[2:].strip().lower())

                row.update({
                    "Guess Rate"    : {"count": float(row_data["Guess Rate"] * 100), "details": [f"{cor}/{tot}"] + self.player_song_details[name]["Overall"]},
                    "UF"            : float (row_data["UF"])                if "UF"     in row_data                 else np.nan,
                    "Score"         : float (row_data["Score"])             if "Score"  in row_data                 else np.nan,
                    "1/8s"          : int   (row_data["1/8s"]),
                    "2/8s"          : int   (row_data["2/8s"]),
                    "7/8s"          : int   (row_data["7/8s"]),
                    "Mean Over-8"   : float (row_data["Mean Over-8"])       if pd.notnull(row_data["Mean Over-8"])  else np.nan
                })

                if self.use_teams: row.update({"Lives Taken": int(row_data["Lives Taken"]), "Lives Saved": int(row_data["Lives Saved"])})
                
                for tid in active: 
                    seen = self.p_type_s[name][tid]
                    succ = self.p_type_c[name][tid]

                    row[t_labels[tid]] = {
                        "count"     : float(row_data[t_labels[tid]] * 100),
                        "details"   : [f"{succ}/{seen}"] + self.player_song_details[name][f"Type {tid}"]
                    } if pd.notnull(row_data[t_labels[tid]]) else np.nan

                if watched:
                    succ_rig    = self.p_rigs_h [name]
                    tot_rig     = self.p_rigs   [name]
                    succ_off    = self.c_counts [name] - succ_rig
                    tot_off     = self.s_part   [name] - tot_rig

                    self.player_song_details[name]["Rigs"].sort(key = lambda s: s[2:].strip().lower())

                    rig_song_lines  = {s[2:] for s in self.player_song_details[name]["Rigs"]}
                    off_details     = [s for s in self.player_song_details[name]["Overall"] if s[2:] not in rig_song_lines]

                    row.update({
                        "Rigs"              : int   (row_data["Rigs"]),
                        "Rig Rate"          : float (row_data["Rig Rate"]       * 100),
                        "Solo Rigs"         : int   (row_data["Solo Rigs"]),
                        "Solo Rig Rate"     : float (row_data["Solo Rig Rate"]  * 100),
                        "Rig Over-8"        : float (row_data["Rig Over-8"])            if pd.notnull(row_data["Rig Over-8"])   else np.nan,
                        "Over-8 Delta"      : float (row_data["Over-8 Delta"])          if pd.notnull(row_data["Over-8 Delta"]) else np.nan,
                        "Rig Guess Rate"    : {"count": float(row_data["Rig Guess Rate"] * 100), "details": [f"{succ_rig}/{tot_rig}"] + self.player_song_details[name]["Rigs"]} if pd.notnull(row_data["Rig Guess Rate"]) else np.nan,
                        "Off Guess Rate"    : {"count": float(row_data["Off Guess Rate"] * 100), "details": [f"{succ_off}/{tot_off}"] + off_details}                            if pd.notnull(row_data["Off Guess Rate"]) else np.nan,
                        "Rig Delta"         : float (row_data["Rig Delta"]      * 100),
                    })

                seen_chan = self.p_chan_s[name]
                succ_chan = self.p_chan_c[name]

                row["Median Time"] = float(row_data["Median Time"]) if pd.notnull(row_data["Median Time"]) else np.nan

                row["Chant Guess Rate"] = {
                    "count"     : float(row_data["Chant Guess Rate"] * 100),
                    "details"   : [f"{succ_chan}/{seen_chan}"] + self.player_song_details[name]["Chant"]
                } if pd.notnull(row_data["Chant Guess Rate"]) else np.nan

            for key in ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Solo Rigs"]:
                if key not in row               : continue
                if key in ["Rigs", "Solo Rigs"] : player_song_details[name][key].sort(key = lambda s: s[2:].strip().lower())
                else                            : player_song_details[name][key].sort(key = str.lower)

                row[key] = {"count": row[key], "details": player_song_details[name][key]}

            rows.append(row)

        df_players = pd.DataFrame(rows)

        if "Guess Rate" in df_players.columns and "Eru" not in self.tour_label:
            if self.val_str == "default":
                if      self.tour_label == "Watched 2+8"                : th_val = "25, 20, 15, 10, 5"
                elif    self.tour_label in ["Watched", "QuagWatched"]   : th_val = "28, 18, 12, 6"
                elif    self.tour_label in ["Usual", "Quagsual"]        : th_val = "28, 19, 8"
                elif    watched                                         : th_val = "28, 18, 12, 6"
                else                                                    : th_val = "28, 19, 8"

            else: th_val = self.val_str

            try     : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
            except  : th = [28.0, 18.0, 12.0, 6.0]

            gv = df_players["Guess Rate"].map(lambda x: x["count"] if isinstance(x, dict) else x).tolist()

            for t in th:
                f_idx = -1

                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i

                if f_idx != -1 and f_idx < len(df_players) - 1: borders.append(int(f_idx))

        desc_cols = [
            "Elo",              "Guess Rate",       "UF",               "Score",
            "1/8s",             "2/8s",             "Lives Taken",      "Lives Saved",
            "OP Guess Rate",    "ED Guess Rate",    "IN Guess Rate",
            "Rigs",             "Rig Rate",         "Solo Rigs",        "Solo Rig Rate",
            "Over-8 Delta",     "Rig Guess Rate",   "Off Guess Rate",
            "Rig Delta",        "Chant Guess Rate"
        ]

        asc_cols    = ["7/8s", "Median Time", "Mean Over-8", "Rig Over-8"]
        int_cols    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs", "Solo Rigs"]
        rate_cols   = ["Guess Rate", "OP Guess Rate", "ED Guess Rate", "IN Guess Rate", "Chant Guess Rate", "Rig Guess Rate", "Off Guess Rate"]
        stats_hl    = {}

        elo_ser     = df_players["Elo"]         .fillna(0.0) if "Elo" in df_players.columns else pd.Series(0.0, index = df_players.index)
        gr_ser      = df_players["Guess Rate"]  .map(lambda x: x["count"] if isinstance(x, dict) else x).fillna(0.0)
        rig_ser     = df_players["Rigs"]        .map(lambda x: x["count"] if isinstance(x, dict) else x).fillna(0.0) if "Rigs" in df_players.columns else pd.Series(0.0, index = df_players.index)

        mask_series = pd.Series(eligibility, index = df_players.index)

        for col in df_players.columns:
            if col in desc_cols or col in asc_cols:
                if col in int_cols or col in rate_cols  : num = df_players[col].map(lambda x: x["count"] if isinstance(x, dict) else x)
                else                                    : num = df_players[col]

                el_num = num[mask_series].dropna() if col in int_cols else num.dropna()

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

                    elif col in ["OP Guess Rate", "ED Guess Rate", "IN Guess Rate", "Chant Guess Rate"]:
                        best_idx    = int(gr_ser.loc[best_b_idx]    .idxmin()) if not best_b_idx    .empty else None
                        worst_idx   = int(gr_ser.loc[worst_b_idx]   .idxmax()) if not worst_b_idx   .empty else None

                    elif col == "Rig Guess Rate":
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

                if      col == "Player" or isinstance(val, dict)                        : row_dict[col] = val
                elif    pd.isnull(val)  or (isinstance(val, float) and np.isnan(val))   : row_dict[col] = "N/A"
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
            team_song_details[tid]["Total 1/8s"].sort(key = str.lower)

            team_rows.append({
                "Team Leader"   : row_data["Team Leader"],
                "Mean Elo"      : float (row_data["Mean Elo"]),
                "Mean GR"       : float (row_data["Mean GR"]),
                "Total 1/8s"    : {"count": int(row_data["Total 1/8s"]), "details": team_song_details[tid]["Total 1/8s"]},
                "Mean Over-8"   : float (row_data["Mean Over-8"]),
                "Rig Synergy"   : float (row_data["Rig Synergy"]),
                "Off Synergy"   : float (row_data["Off Synergy"]),
                "Shared Rigs"   : float (row_data["Shared Rigs"])
            })

        if team_rows:
            df_teams_temp = pd.DataFrame(team_rows)

            for col in df_teams_temp.columns:
                num = df_teams_temp[col].map(lambda x: x["count"]) if col == "Total 1/8s" else df_teams_temp[col]
                desc    = ["Mean Elo", "Mean GR", "Total 1/8s", "Rig Synergy", "Off Synergy", "Shared Rigs"]
                asc     = ["Mean Over-8"]

                if not num.dropna().empty and (col in desc or col in asc):
                    best_val    = num.dropna().min() if col in asc else num.dropna().max()
                    worst_val   = num.dropna().max() if col in asc else num.dropna().min()

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
                if      k in ["Total 1/8s", "Team Leader"]                      : f_dict[k] = v
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

                tier_data[tr].append({
                    "Player"                : p,
                    "Guess Rate"            : {"count": float(round(gen, 2)), "details": [f"{cor}/{tot}"] + player_song_details[p]["Overall"]},
                    "Lives Taken"           : int(atk),
                    "Lives Taken Details"   : player_song_details[p]["Lives Taken"],
                    "Lives Saved"           : float(round(blk, 2)),
                    "Lives Saved Details"   : player_song_details[p]["Lives Saved"],
                    "Contribution Rate"     : {"count": float(round(con, 2)), "details": [f"{int(atk) + int(blk)}/{cor}"]},
                    "Median Time"           : float(round(spd, 2)) if spd is not None and pd.notnull(spd) else None,
                    "Chanting Guess Rate"   : {"count": float(round(chn, 2)), "details": [f"{chc}/{cht}"] + player_song_details[p]["Chant"]}
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

                if self.p_l_corr[name] and pd.notnull(yl) and pd.notnull(yg): arrow_list.append({
                    "acronym"                   : base_node["acronym"],
                    "name"                      : name,
                    "x_start"                   : float(round(np.mean(self.p_l_corr[name]), 2)),
                    "y_start"                   : float(round(r_vint_med,                   2)),
                    "seasonal_vintage_start"    : r_seas,
                    "x_end"                     : base_node["over8"],
                    "y_end"                     : base_node["vintage"],
                    "seasonal_vintage_end"      : p_seas,
                    "rig_gr"                    : base_node["rig_gr"],
                    "gr"                        : base_node["gr"],
                    "rig_rate"                  : base_node["rig_rate"]
                })

        return scatter_list, arrow_list

    def _create_dashboard_html(self, path, use_teams, watched):
        active = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]
        if len(active) <= 1: active = []

        t_labels        = {1: "OP Guess Rate", 2: "ED Guess Rate", 3: "IN Guess Rate"}
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
            "Player"                    : "☆: New player<br>▲/▼: Subbed in/out<br>(X): 0 rigs/corrects in X round(s)",
            "UF"                        : "Usefulness<br>Calculates this player's contribution to their team, scaled by Elo and songs played",
            "Score"                     : "Calculates this player's value (Usefulness) against what's expected from their Elo<br>50 means this player is playing to expectations",
            "Mean Over-8"               : "Average of correct guessers across songs this player/team guessed correctly",
            "Lives Taken"               : "Count of points won against the opposing team<br>Correct guessers exclusively on their team",
            "Lives Saved"               : "Count of blocks achieved against the opposing team<br>Lone correct guesser for their team whilst the opposing team also has correct guesser(s)",
            "Solo Rigs"                 : "Count of songs exclusively from this player's list",
            "Rig Over-8"                : "Average of correct guessers across songs from this player's list",
            "Over-8 Delta"              : "Rig Over-8 - Mean Over-8<br>Calculates the difficulty gap between this player's list and correct guesses",
            "Rig Delta"                 : "100 * (Correct - Rig) / Correct<br>Calculates this player's performance against their own list",
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

        with open(path / "Data.json", "w", encoding = "utf-8") as f: json.dump(data_payload, f, ensure_ascii = False, indent = 4)
        template_dir = self.script_dir / "help" / "template"

        shutil.copy(template_dir / "Dashboard.html",    path / "Dashboard.html")
        shutil.copy(template_dir / "Styles.css",        path / "Styles.css")
        shutil.copy(template_dir / "Script.js",         path / "Script.js")

    def _export_png(self, df, path, fname, title, mask = None, val_str = "default"):
        if not self.browser_path: return
        df = df.reset_index(drop = True)

        desc = [
            "Elo",              "Guess Rate",
            "UF",               "Score",
            "1/8s",             "2/8s",
            "Lives Taken",      "Lives Saved",
            "OP Guess Rate",    "ED Guess Rate",    "IN Guess Rate",
            "Rigs",             "Rig Rate",
            "Solo Rigs",        "Solo Rig Rate",
            "Over-8 Delta",     "Rig Guess Rate",   "Off Guess Rate",
            "Rig Delta",        "Chant Guess Rate",
            "Mean Elo",         "Mean GR",          "Total 1/8s",
            "Rig Synergy",      "Off Synergy",      "Shared Rigs"
        ]

        asc     = ["7/8s", "Median Time", "Mean Over-8", "Rig Over-8"]
        rest    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs"]
        stats   = {}
        elo_col = "Elo" if "Elo" in df.columns else "Mean Elo" if "Mean Elo" in df.columns else None
        elo_ser = pd.to_numeric(df[elo_col], errors = 'coerce').fillna(0.0) if elo_col else pd.Series(0.0, index = df.index)

        for col in df.columns:
            if col in desc or col in asc:
                num     = pd.to_numeric(df[col].astype(str).str.replace('%',''), errors = 'coerce')
                el_num  = num[mask].dropna() if mask is not None and col in rest else num.dropna()

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
                    gr_cols = ["OP Guess Rate", "ED Guess Rate", "IN Guess Rate", "Chant Guess Rate"]
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
                        best_idx    = pd.to_numeric(df["Guess Rate"], errors = 'coerce').fillna(0).loc[best_b_indices]  .idxmin() if not best_b_indices     .empty else None
                        worst_idx   = pd.to_numeric(df["Guess Rate"], errors = 'coerce').fillna(0).loc[worst_b_indices] .idxmax() if not worst_b_indices    .empty else None

                    elif col == "Rig Guess Rate" and "Rigs" in df.columns:
                        best_idx    = rig_ser.loc[best_b_indices]   .idxmax() if not best_b_indices     .empty else None
                        worst_idx   = elo_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                    else:
                        best_idx    = elo_ser.loc[best_b_indices]   .idxmin() if not best_b_indices     .empty else None
                        worst_idx   = elo_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                    stats[col] = {'best_idx': best_idx, 'worst_idx': worst_idx}

        borders = []

        if "Guess Rate" in df.columns:
            if "Eru" in self.tour_label: th = []

            else:
                if val_str == "default":
                    if      self.tour_label == "Watched 2+8"                : th_val = "25, 20, 15, 10, 5"
                    elif    self.tour_label in ["Watched", "QuagWatched"]   : th_val = "28, 18, 12, 6"
                    elif    self.tour_label in ["Usual", "Quagsual"]        : th_val = "28, 19, 8"
                    elif    "Rigs"          in df.columns                   : th_val = "28, 18, 12, 6"
                    else                                                    : th_val = "28, 19, 8"

                else: th_val = val_str

                try     : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
                except  : th = [28.0, 18.0, 12.0, 6.0]

            gv = pd.to_numeric(df["Guess Rate"].astype(str).str.replace('%',''), errors = 'coerce').tolist()

            for t in th:
                f_idx = -1

                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i

                if f_idx != -1 and f_idx < len(df) - 1: borders.append(f_idx)

        col_borders = {"Player", "Guess Rate", "Score", "Mean Over-8", "Lives Saved", "IN Guess Rate", "Rig Rate", "Solo Rig Rate", "Over-8 Delta", "Rig Delta", "Metric", "Value", "Team Leader", "Mean Over-8"}
        th_cells    = []

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
                cnt     =   f"<b>{cell}</b>" if cname in bold_columns else cell
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