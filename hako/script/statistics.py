import json, re
import numpy    as np
import pandas   as pd

from .config        import *
from .integrations  import *
from bs4            import BeautifulSoup
from collections    import defaultdict

def get_threshold_borders(analyzer, df_gr_series) -> list[int]:
    if "Eru" in analyzer.tour_label: return []

    if analyzer.val_str == "default":
        if      analyzer.tour_label == "Watched 2+8s"   : th_val = "25, 20, 15, 10, 5"
        elif    "Watched" in analyzer.tour_label        : th_val = "28, 18, 12, 6"
        else                                            : th_val = "28, 19, 8"
    else                                                : th_val = analyzer.val_str

    try                 : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
    except Exception    : th = [28.0, 18.0, 12.0, 6.0]

    gv      = df_gr_series.tolist()
    borders = []

    for t in th:
        f_idx = -1

        for i, v in enumerate(gv):
            if pd.notnull(v) and v >= t: f_idx = i

        if f_idx != -1 and f_idx < len(df_gr_series) - 1: borders.append(int(f_idx))

    return borders

def compute_player_performance_scores(plist: list, analyzer, elo_map: dict, avg_rank: float) -> list[float]:
    uf_pool, el_pool, gr_pool = [], [], []

    for name in plist:
        tot         = analyzer.s_part[name]
        uf_scaled   = (analyzer.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0
        gr_val      = analyzer.c_counts[name] / tot if tot else 0.0

        try                 : elo = float(elo_map.get(name.lower(), 0.0))
        except Exception    : elo = 0.0

        uf_pool.append(uf_scaled)
        el_pool.append(elo)
        gr_pool.append(gr_val)

    if len(el_pool) > 1 and np.var(el_pool) > 0:
        els = np.array(el_pool)
        ufs = np.array(uf_pool)
        grs = np.array(gr_pool)

        s_uf, i_uf = np.polyfit(els, ufs, 1)
        s_gr, i_gr = np.polyfit(els, grs, 1)

        res_uf = ufs - (s_uf * els + i_uf)
        res_gr = grs - (s_gr * els + i_gr)

        std_uf = np.std(res_uf) if np.std(res_uf) > 0 else 1.0
        std_gr = np.std(res_gr) if np.std(res_gr) > 0 else 1.0

        scores_uf = (1 / (1 + np.exp(SCALE_PERF * (res_uf / std_uf)))) * 100
        scores_gr = (1 / (1 + np.exp(SCALE_PERF * (res_gr / std_gr)))) * 100

        return list((scores_uf * 0.5) + (scores_gr * 0.5))

    return [50.0] * len(plist)

def compute_player_rows(
    analyzer,
    elo_map     : dict,
    _           : dict,
    exp_map     : dict,
    base_exp    : int,
    new_players : list,
    watched     : bool,
    active      : list,
    t_labels    : dict,
    avg_rank    : float,
) -> tuple[pd.DataFrame, list[bool]]:
    rows, eligibility = [], []
    
    plist       = list(analyzer.s_part.keys())
    perf_scores = compute_player_performance_scores(plist, analyzer, elo_map, avg_rank)
    perf_map    = dict(zip(plist, perf_scores))

    for name in analyzer.s_part:
        tot, cor    = analyzer.s_part[name], analyzer.c_counts[name]
        target      = exp_map.get(name, base_exp)
        d_name      = name

        if name in new_players: d_name += " ★"

        if target != "ignore" and target < base_exp:
            if name.lower() in analyzer.main_roster_names   : d_name += " ▼"
            else                                            : d_name += " ▲"

        is_eligible = not ("▼" in d_name or "▲" in d_name)
        eligibility.append(is_eligible)

        history_baselines   = {"GR": np.nan, "UF": np.nan, "OP": np.nan, "ED": np.nan, "IN": np.nan}
        alias_txt_path      = analyzer.tour_dir / FILE_ALIAS

        if alias_txt_path.exists():
            try:
                with open(alias_txt_path, "r", encoding = "utf-8") as f_alias:
                    for a_line in f_alias:
                        if "," in a_line:
                            p_splits = [x.strip() for x in a_line.split(",")]
                            if len(p_splits) >= 7 and (p_splits[0].lower() == name.lower() or p_splits[1].lower() == name.lower()):
                                if p_splits[2] != "N/A": history_baselines = {
                                        "GR": float(p_splits[2]),
                                        "UF": float(p_splits[3]),
                                        "OP": float(p_splits[4]),
                                        "ED": float(p_splits[5]),
                                        "IN": float(p_splits[6]),
                                }

                                break
            except Exception: pass

        row = {"Player": d_name}
        if analyzer.use_teams: row["Elo"] = elo_map.get(name.lower(), np.nan)

        current_gr = cor / tot if tot else 0.0
        row.update({"GR": current_gr})

        delta_gr = (current_gr * 100) - history_baselines["GR"] if pd.notnull(history_baselines["GR"]) else np.nan
        row.update({"GR Δ": round(delta_gr, 2) if pd.notnull(delta_gr) else np.nan})

        if analyzer.use_teams:
            uf_val = (analyzer.p_usefulness_sum[name] * avg_rank * 8) / tot if tot else 0.0
            row.update({"UF": uf_val})

            elo_val = float(elo_map.get(name.lower(), np.nan))

            if pd.notnull(elo_val) and elo_val != 0 : delta_uf = 100 * (uf_val - elo_val) / elo_val
            else                                    : delta_uf = np.nan

            row.update({"UF Δ"  : round(delta_uf, 2) if pd.notnull(delta_uf) else np.nan})
            row.update({"Score" : perf_map.get(name, 50.0)})

        avg_over8 = analyzer.p_overs_sum[name] / cor if cor else np.nan
        row.update({"1/8s": analyzer.e_counts[name], "2/8s": analyzer.p_two_e[name], "7/8s": analyzer.p_rev_e[name], "Mean Over-8": avg_over8})
        if analyzer.use_teams: row.update({"Lives Taken": analyzer.p_pts[name], "Lives Saved": analyzer.p_blks[name]})

        for tid in active:
            seen                = analyzer.p_type_s[name][tid]
            current_type_gr     = analyzer.p_type_c[name][tid] / seen if seen else np.nan
            row[t_labels[tid]]  = current_type_gr
            t_key               = t_labels[tid].split(" ")[0]
            hist_base_val       = history_baselines.get(t_key, np.nan)

            if pd.notnull(current_type_gr) and pd.notnull(hist_base_val):
                delta_type          = (current_type_gr * 100) - hist_base_val
                row[f"{t_key} Δ"]   = round(delta_type, 2)
            else: row[f"{t_key} Δ"] = np.nan

        if watched:
            rig_over8 = np.mean(analyzer.p_l_corr[name]) if analyzer.p_l_corr[name] else np.nan

            row.update(
                {
                    "Rigs"          : analyzer.p_rigs[name],
                    "Rig Rate"      : analyzer.p_rigs[name] / tot                                       if tot                              else np.nan,
                    "Solo Rigs"     : analyzer.p_l_solos[name],                 
                    "Solo Rig Rate" : analyzer.p_l_solos[name] / analyzer.p_rigs[name]                  if analyzer.p_rigs[name]            else np.nan,
                    "Rig Over-8"    : rig_over8,                
                    "Over-8 Δ"      : rig_over8 - avg_over8,                
                    "Rig GR"        : analyzer.p_rigs_h[name] / analyzer.p_rigs[name]                   if analyzer.p_rigs[name]            else np.nan,
                    "Off GR"        : (cor - analyzer.p_rigs_h[name]) / (tot - analyzer.p_rigs[name])   if (tot - analyzer.p_rigs[name])    else np.nan,
                    "Rig Δ"         : (cor - analyzer.p_rigs[name]) / cor                               if cor                              else np.nan,
                }
            )

        h_diffs = analyzer.p_hit_diff.get(name, [])
        h_vints = analyzer.p_hit_vint.get(name, [])

        row.update({"Mean Difficulty Hit": np.mean(h_diffs) if h_diffs else np.nan, "Median Vintage Hit": np.median(h_vints) if h_vints else np.nan})

        times       = analyzer.p_answer_times.get(name, [])
        seen_chan   = analyzer.p_chan_s[name]

        row["Median Time"]  = np.median(times) if times else np.nan
        row["Chant GR"]     = analyzer.p_chan_c[name] / seen_chan if seen_chan else np.nan

        rows.append(row)

    df = pd.DataFrame(rows)

    if "Score" in df.columns: df = df.sort_values(by = ["GR", "Score"], ascending = [False, False])
    elif "Elo" in df.columns:
        df["_sort_elo"] = pd.to_numeric(df["Elo"], errors = "coerce")
        df              = df.sort_values(by = ["GR", "_sort_elo"], ascending = [False, True]).drop(columns = ["_sort_elo"])
    else: df = df.sort_values("GR", ascending = False)

    mask = pd.Series(eligibility, index = pd.DataFrame(rows).index).reindex(df.index).values
    return df, mask

def compute_tour_stats(analyzer, use_teams: bool, watched: bool) -> list:
    def fmt_most(names, val):
        if not names: return "N/A", None

        win = sorted(names, key = lambda x: (analyzer.c_counts[x] / analyzer.s_part[x]) if analyzer.s_part[x] else 0)[0]
        gr  = (analyzer.c_counts[win] / analyzer.s_part[win]) * 100 if analyzer.s_part[win] else 0

        return f"{win} ({val}{f', {gr:.2f}' if len(names) > 1 else ''})", win

    stats = [
        ["Median Vintage",  format_year(round(np.median(analyzer.all_vint), 2))                             if analyzer.all_vint    else "N/A", None],
        ["Mean Difficulty", f"{np.mean(analyzer.all_diff):.2f}"                                             if analyzer.all_diff    else "N/A", None],
        ["Mean GR",         f"{100 * (analyzer.global_stats['tot_c'] / sum(analyzer.s_part.values())):.2f}" if analyzer.s_part      else "N/A", None],

        ["Total 0/8s", analyzer.global_stats["blanks"],     "Total 0/8s"],
        ["Total 1/8s", analyzer.global_stats["solos"],      "Total 1/8s"],
        ["Total 2/8s", analyzer.global_stats["doubles"],    "Total 2/8s"],
        ["Total 7/8s", analyzer.global_stats["sevens"],     "Total 7/8s"],
        ["Total 8/8s", analyzer.global_stats["fulls"],      "Total 8/8s"],
    ]

    if use_teams: stats.append(["Total 4-0s", analyzer.global_stats["sweeps"], "Total 4-0s"])

    pop_gen = analyzer.genre_c  .most_common(1)[0][0] if analyzer.genre_c   else "N/A"
    pop_tag = analyzer.tag_c    .most_common(1)[0][0] if analyzer.tag_c     else "N/A"

    pop_gen_cnt = analyzer.genre_c  .most_common(1)[0][1] if analyzer.genre_c   else 0
    pop_tag_cnt = analyzer.tag_c    .most_common(1)[0][1] if analyzer.tag_c     else 0

    m1_p = [n for n, v in analyzer.e_counts .items() if v == max(analyzer.e_counts  .values(), default = 0) and v > 0]
    m2_p = [n for n, v in analyzer.p_two_e  .items() if v == max(analyzer.p_two_e   .values(), default = 0) and v > 0]
    m7_p = [n for n, v in analyzer.p_rev_e  .items() if v == max(analyzer.p_rev_e   .values(), default = 0) and v > 0]

    f1, w1 = fmt_most(m1_p, max(analyzer.e_counts   .values(), default = 0))
    f2, w2 = fmt_most(m2_p, max(analyzer.p_two_e    .values(), default = 0))
    f7, w7 = fmt_most(m7_p, max(analyzer.p_rev_e    .values(), default = 0))

    stats.extend([
            ["Most Popular Genre",  f"{pop_gen} ({pop_gen_cnt})" if analyzer.genre_c    else "N/A", f"Genre: {pop_gen}"],
            ["Most Popular Tag",    f"{pop_tag} ({pop_tag_cnt})" if analyzer.tag_c      else "N/A", f"Tag: {pop_tag}"],

            ["Most 1/8s",           f1, ("1/8s", w1)],
            ["Most 2/8s",           f2, ("2/8s", w2)],
            ["Most 7/8s",           f7, ("7/8s", w7)]
    ])

    plist   = list(analyzer.s_part.keys())
    no_s    = sorted([n for n in plist if analyzer.e_counts[n] == 0 and analyzer.s_part[n] > 0], key = lambda x: analyzer.c_counts[x] / analyzer.s_part[x], reverse = True)
    yes_s   = sorted([n for n in plist if analyzer.e_counts[n] > 0 and analyzer.s_part[n] > 0], key = lambda x: analyzer.c_counts[x] / analyzer.s_part[x])

    if no_s     : stats.append(["Highest GR Without 1/8s",  f"{no_s[0]}     ({100 * (analyzer.c_counts[no_s     [0]] / analyzer.s_part[no_s     [0]]):.2f})",   None])
    if yes_s    : stats.append(["Lowest GR With 1/8s",      f"{yes_s[0]}    ({100 * (analyzer.c_counts[yes_s    [0]] / analyzer.s_part[yes_s    [0]]):.2f},     {analyzer.e_counts[yes_s[0]]})", ("1/8s", yes_s[0])])

    if watched:
        conv        = []
        eligible    = [p for p in plist if analyzer.p_l_solos[p] > 0]

        if eligible:
            total_hits      = sum((analyzer.p_l_solos[p] - analyzer.p_m_erigs[p]) for p in eligible)
            total_attempts  = sum(analyzer.p_l_solos[p] for p in eligible)
            global_avg      = total_hits / total_attempts if total_attempts > 0 else 0

            for n in eligible:
                t               = analyzer.p_l_solos[n]
                h               = t - analyzer.p_m_erigs[n]
                weighted_score  = (h + CONST_CONV * global_avg) / (t + CONST_CONV)

                conv.append({"n": n, "score": weighted_score, "p": 100 * h / t, "h": h, "t": t})

            b = sorted(conv, key = lambda x: x["score"], reverse = True)    [0]
            w = sorted(conv, key = lambda x: x["score"])                    [0]

            stats.append(["Best Solo Rig Converter",    f"{b['n']} ({b['p']:.2f}, {b['h']}/{b['t']})", ("Solo Rigs", b["n"])])
            stats.append(["Worst Solo Rig Converter",   f"{w['n']} ({w['p']:.2f}, {w['h']}/{w['t']})", ("Solo Rigs", w["n"])])

    return stats

def compute_team_rows(analyzer, assigns: dict, t1_lookup: dict) -> pd.DataFrame:
    if not getattr(analyzer, "use_teams", False): return pd.DataFrame()

    if not t1_lookup:
        for tid, players in analyzer.rosters.items():
            if players:
                sorted_players = sorted(list(players), key=str.lower)
                t1_lookup[tid] = sorted_players[0]

    chal_matches = []

    if getattr(analyzer, "challonge_choice", "No") == "Yes":
        codes_file  = analyzer.tour_dir / FILE_CODES
        chal_link   = None

        if codes_file.exists():
            with open(codes_file, "r", encoding = "utf-8") as f:
                for line in f:
                    if line.strip().startswith("http"):
                        chal_link = line.strip()
                        break

        if chal_link:
            try:
                html = download_challonge_page(chal_link)
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
    alias_path  = analyzer.tour_dir / FILE_ALIAS

    if alias_path.exists():
        with open(alias_path, "r", encoding = "utf-8") as f:
            for line in f:
                if "," in line:
                    parts = [p.strip().lower() for p in line.split(",")]

                    if len(parts) >= 2:
                        alias_map[parts[0]] = parts[1]
                        alias_map[parts[1]] = parts[0]

    res = []

    for tid in analyzer.t_c_ps:
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

            try                 : s1, s2 = int(scores[0]), int(scores[1])
            except Exception    : continue

            p1_leader = parse_challonge_display_leader(m["player1"].get("display_name", ""))
            p2_leader = parse_challonge_display_leader(m["player2"].get("display_name", ""))

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
            "total_matches" : w_count + l_count + t_count,
        }

        t_overs = []

        for original_name in analyzer.s_part:
            if original_name.lower() in assigns:
                if assigns[original_name.lower()][0] == tid and analyzer.c_counts[original_name] > 0:
                    t_overs.append(analyzer.p_overs_sum[original_name] / analyzer.c_counts[original_name])

        t_elos = []

        for p in analyzer.rosters[tid]:
            v = analyzer.elo_map.get(p.lower())

            if v is not None:
                try                 : t_elos.append(float(v))
                except Exception    : pass

        row = {
            "Team Leader"   : leader_name,
            "Mean Elo"      : np.mean(t_elos),
            "Mean GR"       : np.mean(analyzer.t_c_ps[tid])     * 100,
            "Total 1/8s"    : analyzer.t_solos[tid],
            "Mean Over-8"   : np.mean(t_overs),
            "Rig Synergy"   : np.mean(analyzer.t_on_syn[tid])   * 100,
            "Off Synergy"   : np.mean(analyzer.t_off_syn[tid])  * 100,
            "Shared Rigs"   : np.mean(analyzer.t_sh_rig[tid])   * 100,
            "_tid"          : tid,
            "_history"      : h_payload,
        }

        if getattr(analyzer, "text_var_wlt", "No") == "Yes" or h_payload["total_matches"] > 0:
            row["Win Record"]   = f"{w_count}-{l_count}" if t_count == 0 else h_payload["summary"]
            tot                 = w_count + l_count + t_count
            row["_win_pct"]     = (w_count / tot) if tot > 0 else -1.0

        res.append(row)

    df = pd.DataFrame(res).sort_values(by = ["Mean GR", "Mean Elo"], ascending = [False, True])
    return df

def compute_tier_rows(analyzer, assigns: dict, has_chanting_songs: bool) -> tuple[list[dict], list[dict]]:
    rows1, rows2 = [], []

    for tr in ["1", "2", "3", "4"]:
        tp = [n for n in analyzer.s_part if n.lower() in assigns and assigns[n.lower()][1] == tr]
        if not tp: continue

        row1 = {"Tier": tr}
        row2 = {"Tier": tr}

        gen_players, atk_players, blk_players, con_players, spd_players, chn_players = [], [], [], [], [], []

        for p in tp:
            cor = analyzer.c_counts[p]
            tot = analyzer.s_part[p]
            tim = analyzer.p_answer_times.get(p, [])
            chc = analyzer.p_chan_c[p]
            cht = analyzer.p_chan_s[p]

            gen = 100 * cor / tot if tot else 0.0
            atk = analyzer.p_pts[p]
            blk = analyzer.p_blks[p]
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
        row2["Chant GR"]            = f"{chn_players[0]['player']} ({chn_players[0]['value']:.2f})" if chn_players and chn_players[0]["value"] > 0  else ""

        row1["gen_val"] = gen_players[0]["value"] if gen_players else 0.0
        row1["atk_val"] = atk_players[0]["value"] if atk_players else 0.0
        row1["blk_val"] = blk_players[0]["value"] if blk_players else 0.0
        row2["con_val"] = con_players[0]["value"] if con_players else 0.0
        row2["spd_val"] = spd_players[0]["value"] if spd_players else 0.0
        row2["chn_val"] = chn_players[0]["value"] if chn_players else 0.0

        row1["_players"] = {"gen": gen_players, "atk": atk_players, "blk": blk_players}
        row2["_players"] = {"con": con_players, "spd": spd_players, "chn": chn_players}

        rows1.append(row1)
        rows2.append(row2)

    return rows1, rows2