import json, math, os, re
import matplotlib.pyplot    as plt
import matplotlib.colors    as mc
import matplotlib.ticker    as mt
import numpy                as np
import pandas               as pd

from adjustText     import adjust_text
from collections    import Counter, defaultdict
from help.config    import *
from help.dialog    import *
from html2image     import Html2Image
from pathlib        import Path
from PIL            import Image, ImageChops, ImageOps
from scipy.spatial  import ConvexHull
from tkinter        import messagebox

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
        self.p_type_c                   = defaultdict(lambda: defaultdict(int))
        self.p_type_s                   = defaultdict(lambda: defaultdict(int))
        self.p_rigs                     = defaultdict(int)
        self.p_rigs_h                   = defaultdict(int)
        self.p_l_vint                   = defaultdict(list)
        self.p_c_vint                   = defaultdict(list)
        self.p_l_corr                   = defaultdict(list)
        self.p_m_erigs                  = defaultdict(int)
        self.p_l_solos                  = defaultdict(list)
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
        self.chanting_ids               = set()
        self.subbed_players_set         = set()
        self.tour_label                 = ""
        self.id_database                = {}
        self.player_acronyms            = {}

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

    def run(self):
        chanting_path = self.script_dir / DIR_TOURS / FILE_CHANT

        if chanting_path.exists():
            with open(chanting_path, "r") as f:
                for line in f:
                    line = line.strip()
                    if line: self.chanting_ids.add(line)

        json_dir = self.tour_dir / DIR_JSONS

        if not json_dir.exists() or not any(json_dir.glob("*.json")):
            messagebox.showerror("Error", f"Folder not found or empty: {json_dir}")
            return

        json_paths              = list(json_dir.glob("*.json"))
        all_known, appearances  = self._scan_players    (json_paths)

        self._generate_acronyms(all_known)

        use_teams, elo_map, assignments, t1_lookup, rosters, all_known  = self._load_team_data  (all_known)
        missing_list_count                                              = 0
        tour_types                                                      = set()

        for path in json_paths:
            with open(path, encoding = "utf-8") as f: data = json.load(f)
            songs = data.get("songs", [])
            if not songs: continue
            raw_f_players = set()

            for s in songs:
                for p in s.get("correctGuessPlayers", []):
                    if      isinstance(p, str)                  : raw_f_players.add(p)
                    elif    isinstance(p, dict) and "name" in p : raw_f_players.add(p["name"])

                for ls in s.get("listStates", []):
                    if "name" in ls: raw_f_players.add(ls["name"])

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
                    f_type_totals[st] += 1
                    tour_types.add(st)

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

                raw_correct = song.get("correctGuessPlayers", [])
                correct     = set()

                for p in raw_correct:
                    if      isinstance(p, str)                  : correct.add(p)
                    elif    isinstance(p, dict) and "name" in p : correct.add(p["name"])

                self.song_history.append((correct, raw_f_players))

                ls                          =   song.get("listStates", [])
                self.global_stats["tot_c"]  +=  len(correct)
                yr                          =   extract_year(si.get("vintage"))

                if yr is not None                                       : self.all_vint.append(yr)
                if isinstance(si.get("animeDifficulty"), (int, float))  : self.all_diff.append(si.get("animeDifficulty"))
                if not ls                                               : missing_list_count += 1

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
                            self.t_on_syn[tid].append(len(c_on_t)                   / 4.0)
                            self.t_sh_rig[tid].append((len(s_riggers & ros) - 1)    / 3.0)

                        else: self.t_off_syn[tid].append(len(c_on_t)/ 4.0)

                if len(final_members - correct) == 0: self.global_stats["fulls"] += 1

                elif apply_rev and len(final_members - correct) == 1:
                    self.global_stats   ["sevens"]                          += 1
                    self.p_rev_e        [list(final_members - correct)[0]]  += 1

                elif len(correct) == 2:
                    self.global_stats   ["doubles"]     += 1
                    for p in correct: self.p_two_e[p]   += 1

                elif len(correct) == 1:
                    self.global_stats["solos"]  +=  1
                    sw                          =   list(correct)[0]
                    self.e_counts[sw]           +=  1

                    if sw.lower() in assignments:
                        self.t_solos[assignments[sw.lower()][0]] += 1

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
                        self.c_counts[name]     += 1
                        self.p_overs_sum[name]  += len(correct)

                        if st in [1, 2, 3]  : self.p_type_c[name][st]   += 1
                        if is_chan          : self.p_chan_c[name]       += 1
                        if yr is not None   : self.p_c_vint[name].append(yr)

                    if is_chan: self.p_chan_s[name] += 1

                if ls:
                    for p in ls:
                        n = p["name"]
                        self.p_rigs[n] += 1

                        if n in correct     : self.p_rigs_h [n] += 1
                        if yr is not None   : self.p_l_vint [n].append(yr)

                        self.p_l_corr[n].append(len(correct))

        self._finalize_outputs(missing_list_count, appearances, use_teams, elo_map, assignments, t1_lookup, tour_types)

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

        self.main_roster_names                      = set()
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

        def find_best_match(p_in, allow_manual = False, line_text = ""):
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

            if not match and allow_manual and ("[" in line_text or "Subs:" in line_text)    : match = ManualMatchDialog(None, p_in, avail).result
            if match                                                                        : new_aliases[p_in] = match

            return match

        with open(codes, "r", encoding = "utf-8") as f: lines = f.readlines()

        for line in lines:
            matches = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', line)

            for p_in, val in matches:
                match = find_best_match(p_in, allow_manual = True, line_text = line)
                if match: elo_map[match.lower()] = val

        idx                 = 1
        sub_candidates_raw  = []

        for line in lines:
            if "Subs:" in line or "subs:" in line:
                mems_subs = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', line)

                for p_sub, _ in mems_subs:
                    m_sub = find_best_match(p_sub)

                    if m_sub:
                        self.subbed_players_set.add(m_sub.lower())
                        sub_candidates_raw.append(m_sub)

                continue

            if "|" in line  : sec = line.split("|")[0]
            else            : sec = line

            mems = re.findall(r'([^\s(]+)\s*\(([-]?\d+\.\d+)\)', sec)
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

        all_team_ids = set(t1_lookup.keys())

        for sub_player in sub_candidates_raw:
            s_low = sub_player.lower()
            if s_low in assignments: continue
            s_match = next((m for m in assignments if m in s_low or s_low in m), None)

            if s_match  : s_team, s_tier = assignments[s_match]
            else        : s_team, s_tier = (list(all_team_ids)[0] if all_team_ids else 1), "1"

            s_team_name         = self._get_team_acronym(t1_lookup.get(s_team, ""), s_team)
            all_team_names_map  = {self._get_team_acronym(t1_lookup.get(tid, ""), tid): tid for tid in all_team_ids}
            dialog              = SubstitutePromptDialog(None, sub_player, s_team_name, s_tier, all_team_names_map.keys())

            if dialog.result:
                chosen_team_name, chosen_tier   = dialog.result
                chosen_team_id                  = all_team_names_map.get(chosen_team_name, s_team)
                assignments[s_low]              = (chosen_team_id, chosen_tier)
                rosters[chosen_team_id].add(sub_player)

            else:
                assignments[s_low] = (s_team, s_tier)
                rosters[s_team].add(sub_player)

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

        return True, elo_map, assignments, t1_lookup, rosters, all_known

    def _finalize_outputs(self, missing_count, appearances, use_teams, elo_map, assignments, t1_lookup, tour_types):
        watched_valid       = missing_count <= 5
        baseline_initial    = int(np.median([len(appearances.get(name, [])) for name in self.s_part]))

        if len(tour_types) == 1:
            t_map           = {1: "OP", 2: "ED", 3: "IN"}
            t_str           = t_map.get(list(tour_types)[0], "")
            init_label      = f"Watched {t_str}" if watched_valid else f"Random {t_str}"

        else: init_label    = "Watched" if watched_valid else "Usual"

        if "Eru" in init_label and use_teams: default_th = ""

        else:
            if      init_label == "Watched 2+8"                 : default_th = "25, 20, 15, 10, 5"
            elif    init_label in ["Watched",   "QuagWatched"]  : default_th = "28, 18, 12, 6"
            elif    init_label in ["Usual",     "Quagsual"]     : default_th = "28, 19, 8"
            elif    "Rigs" in self.s_part                       : default_th = "28, 18, 12, 6"
            else                                                : default_th = "28, 19, 8"

        has_chanting_songs = any(self.p_chan_s.values())
        if not has_chanting_songs: init_label += " -Chanting"

        meta_dialog     = TourMetadataDialog(None, self.tour_id, init_label, default_th, baseline_initial, list(self.s_part.keys()), elo_map)
        meta_res        = meta_dialog.result if meta_dialog.result else {"tour_label": init_label, "th_str": "default", "base_exp": baseline_initial, "selected_new": []}
        self.tour_label = meta_res["tour_label"]

        if not self.tour_label: self.tour_label = init_label

        val_str     = meta_res["th_str"]
        base_exp    = meta_res["base_exp"]
        new_players = meta_res["selected_new"]

        if "Eru" in self.tour_label and use_teams:
            self.p_pts  .clear()
            self.p_blks .clear()

            for cor, raw_f_players in self.song_history:
                t_list = list({assignments[p.lower()][0] for p in raw_f_players if p.lower() in assignments})

                if len(t_list) == 2:
                    tA, tB  = t_list[0], t_list[1]
                    cA      = {assignments[p.lower()][1]: p for p in raw_f_players if p.lower() in assignments and assignments[p.lower()][0] == tA}
                    cB      = {assignments[p.lower()][1]: p for p in raw_f_players if p.lower() in assignments and assignments[p.lower()][0] == tB}

                    for tr in ["1", "2", "3", "4"]:
                        pA, pB = cA.get(tr), cB.get(tr)

                        if pA and pB:
                            rA, rB = pA in cor, pB in cor

                            if rA and not rB: self.p_pts[pA] += 1
                            if rB and not rA: self.p_pts[pB] += 1
                            if rA and rB:
                                self.p_blks[pA] += 0.50
                                self.p_blks[pB] += 0.50

        t_name              = self.tour_label.strip()
        tour_disp           = f"{t_name} Tour"    
        exp_map             = {}
        mismatched_players  = {}

        for name in list(self.s_part.keys()):
            act = len(appearances.get(name, []))

            if act < base_exp   : mismatched_players[name]  = act
            else                : exp_map[name]             = base_exp

        if mismatched_players:
            mismatch_dialog = MismatchedRoundsDialog(None, mismatched_players, base_exp, self.subbed_players_set)
            mismatch_res    = mismatch_dialog.result if mismatch_dialog.result else {k: base_exp for k in mismatched_players}

            for name, target in mismatch_res.items():
                act             = len(appearances.get(name, []))
                exp_map[name]   = target

                if target > act:
                    avg_songs_per_json  =   sum(self.s_part.values()) / sum(len(v) for v in appearances.values())
                    missing_rounds      =   target - act
                    self.s_part[name]   +=  int(missing_rounds * avg_songs_per_json)

        final_threshold = 6 if len(self.s_part) <= 20 else 5

        if      base_exp >= final_threshold     : stage = "Final"
        elif    base_exp == 3                   : stage = "Mid-Tour"
        else                                    : stage = f"R{base_exp}"

        prefix      = f"{tour_disp}, " 
        out_path    = self.tour_dir / DIR_OUT
        out_path.mkdir(parents = True, exist_ok = True)

        self._create_player_png (use_teams, elo_map, watched_valid, stage, out_path, appearances, prefix, exp_map, base_exp, assignments, new_players, t1_lookup, val_str)
        self._create_tour_png   (use_teams, watched_valid, out_path)

        if watched_valid and assignments    : self._create_team_png     (assignments, t1_lookup, out_path)
        if assignments                      : self._create_tier_png     (assignments, out_path, has_chanting_songs)
        if watched_valid                    : self._create_scatter_png  (out_path, True)

        self._create_scatter_png(out_path)
        if watched_valid: self._create_list_guess_png(out_path)
        self._fuse_and_clean(out_path)

    def _create_player_png(self, use_teams, elo_map, watched, stage, path, apps, prefix, exp_map, base_exp, assigns, new_players, t1_lookup, val_str):
        rows, eligibility   = [], []
        t_labels            = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
        active              = [t for t in [1, 2, 3] if any(self.p_type_s[p][t] > 0 for p in self.s_part)]

        if len(active) <= 1 : active = []

        valid_elos  = [float(v) for v in elo_map.values() if str(v).replace('.', '', 1).isdigit() or (str(v).startswith('-') and str(v)[1:].replace('.', '', 1).isdigit())]
        avg_rank    = np.mean(valid_elos) if valid_elos else 1.0

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
                syms = ["", "(1)", "(2)", "(3)", "(4)", "(5)", "(6)"]
                if 0 < (target-act) < len(syms): d_name += f" {syms[target-act]}"

            row = {"Player": d_name}

            if use_teams:
                team_info = assigns.get(name.lower(), ("N/A", "N/A"))

                if team_info[0] != "N/A":
                    leader_name = t1_lookup.get(team_info[0], "")
                    row["Team"] = self._get_team_acronym(leader_name, team_info[0])
                    row["Tier"] = team_info[1]

                else:
                    row["Team"] = "N/A"
                    row["Tier"] = "N/A"

                row["Elo"] = elo_map.get(name.lower(), "N/A")

            avg_over8 = self.p_overs_sum[name] / cor if cor else np.nan
            row.update({"GR": cor / tot if tot else 0})
            if use_teams: row.update({"UF": (self.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0})
            row.update({"1/8s": self.e_counts[name], "2/8s": self.p_two_e[name], "7/8s": self.p_rev_e[name], "Average Over-8": avg_over8})
            if use_teams: row.update({"Lives Taken": self.p_pts[name], "Lives Saved": self.p_blks[name]})

            for tid in active:
                seen                = self.p_type_s[name][tid]
                row[t_labels[tid]]  = self.p_type_c[name][tid] / seen if seen else np.nan

            if watched:
                rig_over8 = np.mean(self.p_l_corr[name]) if self.p_l_corr[name] else np.nan

                row.update({
                    "Rigs"          : self.p_rigs[name],
                    "Rig Rate"      : self.p_rigs[name]             / tot                       if tot                          else np.nan,
                    "Rig Over-8"    : rig_over8,
                    "Over-8 Delta"  : rig_over8 - avg_over8,
                    "Rig GR"        : self.p_rigs_h[name]           / self.p_rigs[name]         if self.p_rigs[name]            else np.nan,
                    "Off GR"        : (cor - self.p_rigs_h[name])   / (tot - self.p_rigs[name]) if (tot - self.p_rigs[name])    else np.nan,
                    "Rig Delta"     : (cor - self.p_rigs[name])     / cor                       if cor                          else np.nan,
                })

            times       = self.p_answer_times.get(name, [])
            seen_chan   = self.p_chan_s[name]

            row["Median Time"]  = np.median(times) if times else np.nan
            row["Chant GR"]     = self.p_chan_c[name] / seen_chan if seen_chan else np.nan

            rows.append(row)

        df      = pd.DataFrame(rows).sort_values("GR", ascending = False)
        mask    = pd.Series(eligibility, index = pd.DataFrame(rows).index).reindex(df.index).values
        pcts    = ["GR"] + [t_labels[t] for t in active] + (["Rig Rate", "Rig Delta", "Rig GR", "Off GR"] if watched else []) + ["Chant GR"]

        if "Elo"            in df.columns: df["Elo"]            = pd.to_numeric(df["Elo"],              errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "UF"             in df.columns: df["UF"]             = pd.to_numeric(df["UF"],               errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Median Time"    in df.columns: df["Median Time"]    = pd.to_numeric(df["Median Time"],      errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Average Over-8" in df.columns: df["Average Over-8"] = pd.to_numeric(df["Average Over-8"],   errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Rig Over-8"     in df.columns: df["Rig Over-8"]     = pd.to_numeric(df["Rig Over-8"],       errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        if "Over-8 Delta"   in df.columns: df["Over-8 Delta"]   = pd.to_numeric(df["Over-8 Delta"],     errors = 'coerce').map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

        for c in pcts: df[c] = pd.to_numeric(df[c], errors = 'coerce').mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
        self._export_png(df, path, "Player.png", f"{prefix}Player Statistics, {stage}", mask, val_str)

    def _create_tour_png(self, use_teams, watched, path):
        def fmt_most(names, val):
            if not names: return "N/A"

            win = sorted(names, key = lambda x: (self.c_counts[x] / self.s_part[x]) if self.s_part[x] else 0)[0]
            gr  = (self.c_counts[win] / self.s_part[win]) * 100 if self.s_part[win] else 0

            return f"{self._get_player_acronym(win)} ({val}{f', {gr:.2f}' if len(names) > 1 else ''})"

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

        if no_s     : stats.append(["Highest GR Without 1/8s",  f"{self._get_player_acronym(no_s[0])} ({100 * (self.c_counts[no_s[0]] / self.s_part[no_s[0]]):.2f})"])
        if yes_s    : stats.append(["Lowest GR With 1/8s",      f"{self._get_player_acronym(yes_s[0])} ({100 * (self.c_counts[yes_s[0]] / self.s_part[yes_s[0]]):.2f}, {self.e_counts[yes_s[0]]})"])

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

                stats.append(["Best Solo Rig Converter",    f"{self._get_player_acronym(b['n'])} ({b['p']:.2f}, {b['h']}/{b['t']})"])
                stats.append(["Worst Solo Rig Converter",   f"{self._get_player_acronym(w['n'])} ({w['p']:.2f}, {w['h']}/{w['t']})"])

        if use_teams:
            half        = math.ceil(len(stats) / 2)
            left_side   = stats[:half]
            right_side  = stats[half:]

            while len(right_side) < len(left_side): right_side.append(["", ""])
            two_col_stats = []
            for l, r in zip(left_side, right_side): two_col_stats.append([l[0], l[1], r[0], r[1]])
            df_tour = pd.DataFrame(two_col_stats, columns = ["Statistic", "Value", "Statistic", "Value"])

        else: df_tour = pd.DataFrame(stats, columns = ["Statistic", "Value"])

        self._export_png(df_tour, path, "Tour.png", "Tour Statistics")

    def _create_team_png(self, assigns, t1_lookup, path):
        res = []

        for tid in self.t_c_ps:
            leader_name = t1_lookup.get(tid, "")
            t_lbl       = self._get_team_acronym(leader_name, tid)
            t_overs     = []

            for original_name in self.s_part:
                n_lower = original_name.lower()

                if n_lower in assigns:
                    t_info = assigns[n_lower]
                    if t_info[0] == tid and self.c_counts[original_name] > 0: t_overs.append(self.p_overs_sum[original_name] / self.c_counts[original_name])
                    
            avg_o = np.mean(t_overs) if t_overs else np.nan

            res.append({
                "T1"                : t_lbl,
                "Median Vintage"    : format_year(np.median(self.t_vint[tid])),
                "Average GR"        : f"{np.mean(self.t_c_ps    [tid]) * 100:.2f}",
                "Rig Synergy"       : f"{np.mean(self.t_on_syn  [tid]) * 100:.2f}",
                "Off Synergy"       : f"{np.mean(self.t_off_syn [tid]) * 100:.2f}",
                "Shared Rigs"       : f"{np.mean(self.t_sh_rig  [tid]) * 100:.2f}",
                "Average Over-8"    : f"{avg_o:.2f}" if not np.isnan(avg_o) else "N/A",
                "Total 1/8s"        : self.t_solos[tid],
            })

        self._export_png(pd.DataFrame(res).sort_values("Average GR", ascending = False), path, "Team.png", "Team Statistics")

    def _create_tier_png(self, assigns, path, has_chanting_songs):
        categories = ["Generalist", "Attacker", "Blocker", "Contributor", "Speedster"]
        if has_chanting_songs: categories.append("Chanter")

        data_by_cat = {cat: [] for cat in categories}

        for tr in ["1", "2", "3", "4"]:
            tp = [n for n in self.s_part if n.lower() in assigns and assigns[n.lower()][1] == tr]
            if not tp: continue

            gen_players = []

            for p in tp:
                val = 100 * (self.c_counts[p] / self.s_part[p]) if self.s_part[p] else 0
                gen_players.append({"player": self._get_player_acronym(p), "value": val, "tier": tr})

            gen_players.sort(key = lambda x: x["value"], reverse = True)
            data_by_cat["Generalist"].extend(gen_players)

            atk_players = []

            for p in tp:
                val = self.p_pts[p]
                atk_players.append({"player": self._get_player_acronym(p), "value": val, "tier": tr})

            atk_players.sort(key = lambda x: x["value"], reverse = True)
            data_by_cat["Attacker"].extend(atk_players)

            blk_players = []

            for p in tp:
                val = self.p_blks[p]
                blk_players.append({"player": self._get_player_acronym(p), "value": val, "tier": tr})

            blk_players.sort(key = lambda x: x["value"], reverse = True)
            data_by_cat["Blocker"].extend(blk_players)

            con_players = []

            for p in tp:
                val = self.p_pts[p] + self.p_blks[p]
                con_players.append({"player": self._get_player_acronym(p), "value": val, "tier": tr})

            con_players.sort(key = lambda x: x["value"], reverse = True)
            data_by_cat["Contributor"].extend(con_players)

            spd_players = []

            for p in tp:
                times = self.p_answer_times.get(p, [])

                if times:
                    val = np.median(times)
                    spd_players.append({"player": self._get_player_acronym(p), "value": val, "tier": tr})

            spd_players.sort(key = lambda x: x["value"], reverse = False)
            data_by_cat["Speedster"].extend(spd_players)

            if has_chanting_songs:
                chn_players = []

                for p in tp:
                    if self.p_chan_s[p] > 0 and self.c_counts[p] > 0:
                        val = 100 * self.p_chan_c[p] / self.p_chan_s[p]
                        chn_players.append({"player": self._get_player_acronym(p), "value": val, "tier": tr})

                chn_players.sort(key = lambda x: x["value"], reverse = True)
                data_by_cat["Chanter"].extend(chn_players)

        num_plots       = len(categories)
        fig, axes       = plt.subplots(2, 3, figsize = (15, 10))
        axes            = axes.flatten()
        segment_width   = 1.00
        tier_gap        = 1.00

        for idx, cat in enumerate(categories):
            ax      = axes[idx]
            items   = data_by_cat[cat]

            ax.set_title(cat, fontsize = 15, weight = 'bold', fontname = "Segoe UI", pad = 10)

            if not items:
                ax.text(0.5, 0.5, "No Data", ha = 'center', va = 'center', fontsize = 17.5, fontname = "Segoe UI")
                ax.tick_params(axis = 'both', which = 'both', length = 0, labelbottom = False, labelleft = False)
                continue

            if cat == "Chanter":
                ax.set_xlim(0, 100)
                ax.xaxis.set_major_locator(mt.MultipleLocator(20))

            elif cat == "Speedster":
                ax.set_xlim(0, 20)
                ax.xaxis.set_major_locator(mt.MultipleLocator(4))

            else:
                factor  = 10 if cat == "Generalist" else 5
                max_v   = max([item["value"] for item in items]) + 1 if items else factor
                xmax    = min(math.ceil(max_v / factor) * factor, 100)

                if      xmax == 0   : xmax      = factor
                elif    xmax <= 20  : factor    = xmax / 5

                ax.set_xlim(0, xmax)
                ax.xaxis.set_major_locator(mt.MultipleLocator(factor))

            xmin_axis       = 0.0
            y_ticks         = []
            labels          = []
            tier_groups     = defaultdict(list)

            for item in items: tier_groups[item["tier"]].append(item)

            current_y       = 0.0
            sorted_tiers    = sorted(list(tier_groups.keys()))
            first_y         = None
            last_y          = None

            for t in sorted_tiers:
                group       = tier_groups[t]
                num_players = len(group)
                block_start = current_y
                block_end   = block_start + (num_players * segment_width)

                if first_y is None: first_y = block_start
                last_y = block_end

                vertices = [(xmin_axis, block_start)]

                for p_idx, item in enumerate(group):
                    p_start = block_start + (p_idx * segment_width)
                    p_end   = p_start + segment_width

                    y_ticks .append(p_start + (segment_width / 2))
                    labels  .append(item["player"])

                    vertices.append((item["value"], p_start))
                    vertices.append((item["value"], p_end))

                vertices.append((xmin_axis, block_end))

                v_x, v_y = zip(*vertices)
                ax.plot(v_x, v_y, color = 'black', linewidth = 1, zorder = 3)
                current_y = block_end + tier_gap

            ax.set_yticks       (y_ticks)
            ax.set_yticklabels  (labels)
            ax.set_ylim         (first_y, last_y)
            ax.invert_yaxis     ()
            ax.tick_params      (axis = 'both', which = 'both', length = 0, pad = 5, labelsize = 10)

            for label in ax.get_xticklabels(): label.set_fontname("Segoe UI")
            for label in ax.get_yticklabels(): label.set_fontname("Segoe UI")

            ax.grid(False)

        for j in range(num_plots, len(axes)): axes[j].axis('off')

        plt.suptitle        ("Tier Statistics", fontname = "Segoe UI", fontsize = 25, weight = 'bold')
        plt.tight_layout    ()
        plt.savefig         (path / "Tier.png", dpi = 500)
        plt.close           (fig)

        try     : trim_whitespace(path / "Tier.png")
        except  : pass

    def _create_scatter_png(self, path, list_mode = False):
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

                    scale_l = 1.00 if len(plist_l) <= 20 else (0.75 if len(plist_l) <= 28 else 0.50)
                    sizes_l = [rate ** 2 * 10000 * scale_l for rate in rig_rates]
                    cmap_l  = mc.LinearSegmentedColormap.from_list("rig_gr_cmap", [(0.0, "#D95400"), (0.5, "#D95400"), (RIG_GR, "#FFFFFF"), (1.0, "#0056B3")])
                    
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
                        "cbar_ticks"        : [0.0, 0.5, RIG_GR, 1.0],
                        "cbar_ticklabels"   : ['0', '50', f'{int(RIG_GR * 100)}', '100']
                    })

        plist_g = [n for n in self.s_part if self.c_counts[n] > 0]

        if plist_g:
            x_vals_g    = [self.p_overs_sum[name] / self.c_counts[name] for name in plist_g]
            y_vals_g    = [np.median(self.p_c_vint[name]) if self.p_c_vint[name] else np.nan for name in plist_g]
            valid_g     = [(p, x, y) for p, x, y in zip(plist_g, x_vals_g, y_vals_g) if not np.isnan(y)]
            
            if valid_g:
                plist_g, x_vals_g, y_vals_g = zip(*valid_g)
                plist_g, x_vals_g, y_vals_g = list(plist_g), list(x_vals_g), list(y_vals_g)
                
                gr_vals     = [self.c_counts            [name] / self.s_part[name] if self.s_part[name] else 0 for name in plist_g]
                uf_rates    = [self.p_usefulness_sum    [name] / self.s_part[name] if self.s_part[name] else 0 for name in plist_g]

                scale_g = 1.00 if len(plist_g) <= 20 else (0.75 if len(plist_g) <= 28 else 0.50)
                sizes_g = [rate ** 2 * 25000 * scale_g for rate in uf_rates]
                cmap_g  = mc.LinearSegmentedColormap.from_list("guess_gr_cmap", [(0.0, "#D95400"), (GEN_GR * 2, "#FFFFFF"), (1.0, "#0056B3")])
                
                configs.append({
                    "filename"          : "Guess.png",
                    "title"             : "Guess Statistics",
                    "plist"             : plist_g,
                    "x_vals"            : x_vals_g,
                    "y_vals"            : y_vals_g,
                    "sizes"             : sizes_g,
                    "colors"            : gr_vals,
                    "cmap"              : cmap_g,
                    "vmin"              : 0.0,
                    "vmax"              : 0.5,
                    "cbar_label"        : "Guess Rate",
                    "cbar_ticks"        : [0.0, GEN_GR, 0.50],
                    "cbar_ticklabels"   : ['0', f'{int(GEN_GR * 100)}', '50']
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
                step = r // 4
                break

            elif r % 3 == 0:
                step = r // 3
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
            center_of_mass  = np.median(norm_points, axis=0)
            distances       = np.linalg.norm(norm_points - center_of_mass, axis = 1)
            pack_mask       = distances < np.percentile(distances, 90)
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

            ax.set_xticks(np.arange (x_min, x_max + 0.5,    0.5))
            ax.set_yticks(range     (y_min, y_max + 1,      step))

            texts = []

            for name, x, y in zip(cfg["plist"], cfg["x_vals"], cfg["y_vals"]):
                label = self._get_player_acronym(name)
                if not label: continue

                ha_align = "left"   if x >= x_center else "right"
                va_align = "bottom" if y >= y_center else "top"

                texts.append(ax.text(x, y, label, fontsize = 10, fontname = "Segoe UI", ha = ha_align, va = va_align))

            if texts:
                adjust_text(
                    texts,
                    ax                      = ax,
                    objects                 = sc,
                    avoid_self              = True,
                    add_objects_to_edges    = True,
                    force_text              = (1.00, 1.00),
                    force_objects           = (1.00, 1.00),
                    expand                  = (2.00, 2.00),
                    arrowprops              = dict(arrowstyle="-", color='black', shrinkA=10)
                )

            ax.set_title    (cfg["title"],      weight = 'bold', fontname = "Segoe UI", fontsize = 22.5, pad        = 12.5)
            ax.set_xlabel   ("Average Over-8",  weight = 'bold', fontname = "Segoe UI", fontsize = 15.0, labelpad   = 2.5)
            ax.set_ylabel   ("Median Vintage",  weight = 'bold', fontname = "Segoe UI", fontsize = 15.0, labelpad   = 2.5)

            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, _: str(int(val))))
            plt.setp(ax.get_yticklabels(), horizontalalignment = 'center', verticalalignment = 'center')

            ax.tick_params(axis = 'x', which = 'both', length = 0, pad = 5)
            ax.tick_params(axis = 'y', which = 'both', length = 0, pad = 15)

            cbar = fig.colorbar(sc, ax = ax, pad = 0.005, aspect = 40, ticks = cfg["cbar_ticks"])
            cbar.set_label(cfg["cbar_label"], weight = 'bold', fontname = "Segoe UI", fontsize = 15, labelpad = -5)

            cbar.ax.set_yticklabels(cfg["cbar_ticklabels"])
            cbar.ax.tick_params(labelsize = 10, length = 0)

            ax.text(0.01, 0.99, "New\nHard", transform = ax.transAxes, color = "grey", fontsize = 10, va = "top",       ha = "left",    weight = "bold", alpha = 0.75)
            ax.text(0.99, 0.99, "New\nEasy", transform = ax.transAxes, color = "grey", fontsize = 10, va = "top",       ha = "right",   weight = "bold", alpha = 0.75)
            ax.text(0.01, 0.01, "Old\nHard", transform = ax.transAxes, color = "grey", fontsize = 10, va = "bottom",    ha = "left",    weight = "bold", alpha = 0.75)
            ax.text(0.99, 0.01, "Old\nEasy", transform = ax.transAxes, color = "grey", fontsize = 10, va = "bottom",    ha = "right",   weight = "bold", alpha = 0.75)

            ax.grid(False)

            plt.tight_layout    ()
            plt.savefig         (path / cfg["filename"], dpi = 500)
            plt.close           (fig)

            try:
                with Image.open(path / cfg["filename"]) as img:

                    img     = img           .convert    ("RGB")
                    bg      = Image         .new        (img.mode, img.size, "white")
                    diff    = ImageChops    .difference (img, bg)
                    bbox    = diff          .getbbox    ()

                    if bbox:
                        img = img.crop(bbox)
                        img = ImageOps.expand(img, border = 30, fill = "white")

                        img.save(path / cfg["filename"], compress_level = 9, optimize = True)

            except Exception: pass

    def _create_list_guess_png(self, path):
        plist               = []
        x_start, y_start    = [], []
        x_end, y_end        = [], []
        gr_vals, grid_grs   = [], []

        for name in self.s_part:
            if self.p_l_corr[name] and self.c_counts[name] > 0:
                yl = np.median(self.p_l_vint[name]) if self.p_l_vint[name] else np.nan
                yg = np.median(self.p_c_vint[name]) if self.p_c_vint[name] else np.nan

                if not np.isnan(yl) and not np.isnan(yg):
                    plist.append(name)

                    x_start.append(np.mean(self.p_l_corr[name]))
                    y_start.append(yl)

                    x_end.append(self.p_overs_sum[name] / self.c_counts[name])
                    y_end.append(yg)

                    gr_vals     .append(self.c_counts[name] / self.s_part[name] if self.s_part[name] else 0.0)
                    grid_grs    .append(self.p_rigs_h[name] / self.p_rigs[name] if self.p_rigs[name] else 0.0)

        if not plist: return

        all_x = x_start + x_end
        all_y = y_start + y_end

        x_min = math.floor  ((min   (all_x) - 0.5) * 2) / 2
        x_max = math.ceil   ((max   (all_x) + 0.5) * 2) / 2
        y_min = math.floor  (min    (all_y) - 1.0)
        y_max = math.ceil   (max    (all_y) + 1.0)

        while True:
            r = y_max - y_min

            if r % 4 == 0:
                step = r // 4
                break

            elif r % 3 == 0:
                step = r // 3
                break

            if r % 2 != 0 or y_max >= 2026  : y_min -= 1
            else                            : y_max += 1

        fig, ax = plt.subplots(figsize = (10, 10))

        ax.set_xlim(x_min, x_max)
        ax.set_ylim(y_min, y_max)

        ax.set_xticks(np.arange (x_min, x_max + 0.5,    0.5))
        ax.set_yticks(range     (y_min, y_max + 1,      step))

        cmap_l  = mc.LinearSegmentedColormap.from_list("rig_gr_cmap", [(0.0, "#A6611A"), (0.5, "#A6611A"), (RIG_GR, "#E7D4B8"), (1.0, "#018571")])
        norm    = mc.Normalize(vmin = 0.0, vmax = 1.0)

        for name, xl, yl, xg, yg, gr, rig_gr in zip(plist, x_start, y_start, x_end, y_end, gr_vals, grid_grs):
            label = self._get_player_acronym(name)
            if not label: continue

            line_thickness  = (min(1, gr + 0.5)) ** 2 * 2
            arrow_color     = cmap_l(norm(rig_gr))

            ax.annotate(
                "", 
                xy          = (xg, yg), 
                xytext      = (xl, yl),
                arrowprops  = dict(arrowstyle = "->", color = arrow_color, linewidth = line_thickness),
                zorder      = 3
            )

            xm = (xl + xg) / 2
            ym = (yl + yg) / 2

            trans   = ax.transData.transform
            p_start = trans((xl, yl))
            p_end   = trans((xg, yg))
            
            dx = p_end[0] - p_start[0]
            dy = p_end[1] - p_start[1]
            
            angle = np.degrees(np.arctan2(dy, dx))

            if      angle > 90  : angle -= 180
            elif    angle < -90 : angle += 180

            gap         = 2.5
            angle_rad   = np.radians(angle)

            p_mid           = trans((xm, ym))
            p_mid_shifted   = (p_mid[0] - gap * np.sin(angle_rad), p_mid[1] + gap * np.cos(angle_rad))

            xm_shifted, ym_shifted  = ax.transData.inverted().transform(p_mid_shifted)

            ax.text(
                xm_shifted, ym_shifted, label,
                fontsize        = 10,
                fontname        = "Segoe UI",
                weight          = "bold",
                ha              = "center",
                va              = "bottom",
                rotation        = angle,
                rotation_mode   = "anchor",
                zorder          = 4,
            )

        ax.set_title    ("List-Guess Statistics",   weight = 'bold', fontname = "Segoe UI", fontsize = 22.5, pad        = 12.5)
        ax.set_xlabel   ("Average Over-8",          weight = 'bold', fontname = "Segoe UI", fontsize = 15.0, labelpad   = 2.5)
        ax.set_ylabel   ("Median Vintage",          weight = 'bold', fontname = "Segoe UI", fontsize = 15.0, labelpad   = 2.5)

        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, _: str(int(val))))
        plt.setp(ax.get_yticklabels(), horizontalalignment = 'center', verticalalignment = 'center')

        ax.tick_params(axis = 'x', which = 'both', length = 0, pad = 5)
        ax.tick_params(axis = 'y', which = 'both', length = 0, pad = 15)

        sm = plt.cm.ScalarMappable(cmap = cmap_l, norm = norm)
        sm.set_array([])

        cbar = fig.colorbar(sm, ax = ax, pad = 0.005, aspect = 40, ticks = [0.0, 0.5, RIG_GR, 1.0])
        cbar.set_label("Rig GR", weight = 'bold', fontname = "Segoe UI", fontsize = 15, labelpad = -5)
        cbar.ax.set_yticklabels(['0', '50', f'{int(RIG_GR * 100)}', '100'])
        cbar.ax.tick_params(labelsize = 10, length = 0)

        ax.text(0.01, 0.99, "New\nHard", transform = ax.transAxes, color = "grey", fontsize = 10, va = "top",       ha = "left",    weight = "bold", alpha = 0.75)
        ax.text(0.99, 0.99, "New\nEasy", transform = ax.transAxes, color = "grey", fontsize = 10, va = "top",       ha = "right",   weight = "bold", alpha = 0.75)
        ax.text(0.01, 0.01, "Old\nHard", transform = ax.transAxes, color = "grey", fontsize = 10, va = "bottom",    ha = "left",    weight = "bold", alpha = 0.75)
        ax.text(0.99, 0.01, "Old\nEasy", transform = ax.transAxes, color = "grey", fontsize = 10, va = "bottom",    ha = "right",   weight = "bold", alpha = 0.75)

        ax.grid(False)

        plt.tight_layout    ()
        plt.savefig         (path / "List-Guess.png", dpi = 500)
        plt.close           (fig)

        try:
            with Image.open(path / "List-Guess.png") as img:
                img     = img           .convert    ("RGB")
                bg      = Image         .new        (img.mode, img.size, "white")
                diff    = ImageChops    .difference (img, bg)
                bbox    = diff          .getbbox    ()

                if bbox:
                    img = img.crop(bbox)
                    img = ImageOps.expand(img, border = 30, fill = "white")
                    img.save(path / "List-Guess.png", compress_level = 9, optimize = True)

        except Exception: pass

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
            w_cell = "N/A"

            if i < len(best):
                p       = best[i]
                b_cell  = f"{self._get_player_acronym(p)} ({get_ratio(p):.2f})"

            if i < len(worst):
                p       = worst[i]
                w_cell  = f"{self._get_player_acronym(p)} ({get_ratio(p):.2f})"

            rows.append([f"{i + 1}", b_cell, w_cell])

        self._export_png(pd.DataFrame(rows, columns = ["Rank", "Best", "Worst"]), path, "Chanting.png", "Chanting Statistics")

    def _create_time_png(self, path):
        plist = [n for n in self.s_part if len(self.p_answer_times.get(n, [])) > 0]
        if not plist: return

        def get_med(p): return np.median(self.p_answer_times[p])

        fastest = sorted(plist, key = get_med)[ : 3]
        slowest = sorted(plist, key = get_med, reverse = True)[ : 3]
        rows    = []

        for i in range(3):
            f_cell = "N/A"
            s_cell = "N/A"

            if i < len(fastest):
                p       = fastest[i]
                f_cell  = f"{self._get_player_acronym(p)} ({get_med(p):.2f})"

            if i < len(slowest):
                p       = slowest[i]
                s_cell  = f"{self._get_player_acronym(p)} ({get_med(p):.2f})"

            rows.append([f"{i + 1}", f_cell, s_cell])

        self._export_png(pd.DataFrame(rows, columns = ["Rank", "Fastest", "Slowest"]), path, "Time.png", "Time Statistics")

    def _export_png(self, df, path, fname, title, mask = None, val_str = "default"):
        if not self.browser_path: return

        desc = [
            "Elo",          "GR",           "UF",           "1/8s",         "2/8s",
            "Lives Taken",  "Lives Saved",  "OP GR",        "ED GR",        "IN GR",
            "Rigs",         "Rig Rate",     "Over-8 Delta", "Rig GR",       "Off GR",       "Rig Delta", 
            "Chant GR",     "Average GR",   "Rig Synergy",  "Off Synergy",  "Shared Rigs",  "Total 1/8s"
        ]

        asc     = ["7/8s", "Median Time", "Average Over-8", "Rig Over-8"]
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

        if "GR" in df.columns:
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

            gv = pd.to_numeric(df["GR"].astype(str).str.replace('%',''), errors = 'coerce').tolist()

            for t in th:
                f_idx = -1

                for i, v in enumerate(gv):
                    if pd.notnull(v) and v >= t: f_idx = i

                if f_idx != -1 and f_idx < len(df) - 1: borders.append(f_idx)

        col_borders = {"Player", "Tier", "UF", "Average Over-8", "Lives Saved", "IN GR", "Rig Rate", "Over-8 Delta", "Rig Delta", "Statistic", "Value", "T1", "Median Vintage"}
        th_cells    = []

        for c in df.columns:
            s_th = ' style="border-right: 3px solid black;"' if c in col_borders else ''
            th_cells.append(f"<th{s_th}>{str(c).replace(' ', '<br>')}</th>")

        html            = "<thead><tr>" + "".join(th_cells) + "</tr></thead><tbody>"
        bold_columns    = {"Player", "Statistic", "T1"}

        for idx, row in df.iterrows():
            b_s     =   "border-bottom: 3px solid black;" if idx in borders else ""
            html    +=  "<tr>"

            for i, (cname, cell) in enumerate(row.items()):
                style = [b_s] if b_s else []
                if cname in col_borders: style.append("border-right: 3px solid black;")

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
                        font-size           : 30px;
                        text-align          : center
                    }} 
                    table {{
                        border-collapse     : collapse;
                        width               : auto;
                        border              : 3px solid black
                    }} 
                    th {{
                        font-weight         : bold;
                        font-size           : 20px;
                        text-align          : center;
                        padding             : 10px;
                        border              : 1px solid black;
                        border-bottom       : 3px solid black;
                        background-color    : #f0f0f0
                    }} 
                    td {{
                        font-size           : 20px;
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

    def _fuse_and_clean(self, path):
        f       = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "List": "List.png", "Guess": "Guess.png"}
        ps      = {k: path / v for k, v in f.items() if (path / v).exists()}
        imgs    = {k: Image.open(v) for k, v in ps.items()}

        if "Tour" not in imgs:
            for k, p in ps.items():
                if k not in ["List", "Guess"]:
                    try     : os.remove(p)
                    except  : pass

            return

        img_tour    = imgs["Tour"]
        img_team    = imgs.get("Team")
        img_tier    = imgs.get("Tier")
        img_list    = imgs.get("List")
        img_guess   = imgs.get("Guess")

        col1_w = img_tour.width
        col1_h = img_tour.height

        if img_team:
            col1_w = max(col1_w, img_team.width)
            col1_h += 10 + img_team.height

        block1_h = col1_h

        if img_guess:
            guess_aspect        = img_guess.width / img_guess.height
            guess_w_scaled      = int(block1_h * guess_aspect)
            img_guess_scaled    = img_guess.resize((guess_w_scaled, block1_h), Image.Resampling.LANCZOS)
            tour_x_offset       = guess_w_scaled + 10

        else:
            img_guess_scaled    = None
            tour_x_offset       = 0

        block1_w = tour_x_offset + col1_w

        if img_tier:
            tier_aspect     = img_tier.width / img_tier.height
            tier_w_scaled   = int(block1_h * tier_aspect)
            img_tier_scaled = img_tier.resize((tier_w_scaled, block1_h), Image.Resampling.LANCZOS)
            tier_x_offset   = block1_w + 10
            block1_w        = tier_x_offset + img_tier_scaled.width

        else:
            img_tier_scaled = None
            tier_x_offset   = block1_w

        if img_list:
            list_aspect     = img_list.width / img_list.height
            list_w_scaled   = int(block1_h * list_aspect)
            img_list_scaled = img_list.resize((list_w_scaled, block1_h), Image.Resampling.LANCZOS)
            extra_w         = block1_w + 10 + img_list_scaled.width

        else:
            img_list_scaled = None
            extra_w         = block1_w

        extra_h     = block1_h
        extra_img   = Image.new("RGB", (extra_w, extra_h), "white")

        if img_guess_scaled: extra_img.paste(img_guess_scaled, (0, 0))
        extra_img.paste(img_tour, (tour_x_offset, 0))

        if img_team         : extra_img.paste(img_team,         (tour_x_offset, img_tour.height + 10))
        if img_tier_scaled  : extra_img.paste(img_tier_scaled,  (tier_x_offset, 0))
        if img_list_scaled  : extra_img.paste(img_list_scaled,  (block1_w + 10, 0))

        extra_out_p = path / "Extra.png"
        extra_img.save(extra_out_p, compress_level = 9, optimize = True)

        try     : trim_whitespace(extra_out_p)
        except  : pass

        p_path = path / "Player.png"

        if p_path.exists():
            extra_img   = Image.open(extra_out_p)
            img_player  = Image.open(p_path)

            player_aspect       = img_player.width / img_player.height
            player_w_scaled     = extra_img.width
            player_h_scaled     = int(player_w_scaled / player_aspect)
            img_player_scaled   = img_player.resize((player_w_scaled, player_h_scaled), Image.Resampling.LANCZOS)

            gen_w = extra_img.width
            gen_h = img_player_scaled.height + 10 + extra_img.height

            general_img = Image.new("RGB", (gen_w, gen_h), "white")
            general_img.paste(img_player_scaled, (0, 0))
            general_img.paste(extra_img, (0, img_player_scaled.height + 10))

            gen_out_p = path / "General.png"
            general_img.save(gen_out_p, compress_level = 9, optimize = True)

            try     : trim_whitespace(gen_out_p)
            except  : pass