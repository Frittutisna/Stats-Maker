import json, shutil, time
import numpy    as np
import pandas   as pd

from .config        import *
from .statistics    import *
from collections    import defaultdict
from pathlib        import Path

def render_dashboard_player(
    analyzer,
    sorted_players      : list[str],
    active              : list,
    t_labels            : dict,
    watched             : bool,
    player_song_details : dict,
    df_base             : pd.DataFrame,
) -> tuple[list[dict], dict, list[int], list[bool]]:
    rows, eligibility, borders = [], [], []

    for name in sorted_players:
        row_data    = df_base.loc[df_base["Player"].str.startswith(name)].iloc[0] if any(df_base["Player"].str.startswith(name)) else None
        target      = analyzer.exp_map.get(name, analyzer.base_exp)
        d_name      = name
        sub_hover   = ""

        if name in analyzer.new_players: d_name += " ★"

        if target != "ignore" and target < analyzer.base_exp:
            if name.lower() in analyzer.main_roster_names:
                d_name  +=  " ▼"
                subs    =   analyzer.sub_relations.get(name.casefold(), [])
                if subs: sub_hover = f"Subbed by {', '.join(subs)}"

            else:
                d_name  +=  " ▲"
                orig    =   analyzer.sub_relations.get(name.lower(), [])
                if orig: sub_hover = f"Subbing for {orig[0]}"

        is_eligible = not ("▼" in d_name or "▲" in d_name)
        eligibility.append(is_eligible)

        row = {"Player": {"count": d_name, "details": [sub_hover] if sub_hover else []}}

        if analyzer.use_teams:
            t_info      = analyzer.assignments.get(name.lower(), (None, "N/A"))
            team_leader = analyzer.t1_lookup.get(t_info[0], "N/A") if t_info[0] is not None else "N/A"
            row["Team"] = team_leader
            row["Tier"] = t_info[1]

            try                 : row["Elo"] = float(analyzer.elo_map.get(name.lower(), np.nan))
            except Exception    : row["Elo"] = np.nan

        if row_data is not None:
            tot, cor = analyzer.s_part[name], analyzer.c_counts[name]
            for key_details in ["Overall", "Type 1", "Type 2", "Type 3", "Chant"]: analyzer.player_song_details[name][key_details].sort(key = lambda s: s[2:].strip().lower())

            row.update({
                "GR"            : {"count": float(row_data["GR"] * 100), "details": [f"{cor}/{tot}"] + analyzer.player_song_details[name]["Overall"]},
                "GR Δ"          : float (row_data["GR Δ"])          if pd.notnull(row_data.get("GR Δ"))     else np.nan,
                "UF"            : float (row_data["UF"])            if "UF"     in row_data                 else np.nan,
                "UF Δ"          : float (row_data["UF Δ"])          if pd.notnull(row_data.get("UF Δ"))     else np.nan,
                "Score"         : float (row_data["Score"])         if "Score"  in row_data                 else np.nan,
                "1/8s"          : int   (row_data["1/8s"]),
                "2/8s"          : int   (row_data["2/8s"]),
                "7/8s"          : int   (row_data["7/8s"]),
                "Mean Over-8"   : float (row_data["Mean Over-8"])   if pd.notnull(row_data["Mean Over-8"])  else np.nan,
            })

            if analyzer.use_teams: row.update({"Lives Taken": int(row_data["Lives Taken"]), "Lives Saved": int(row_data["Lives Saved"])})

            for tid in active:
                seen        = analyzer.p_type_s[name][tid]
                succ        = analyzer.p_type_c[name][tid]
                t_key       = t_labels[tid].split(" ")[0]
                delta_key   = f"{t_key} Δ"

                row[t_labels[tid]]  = {"count": float(row_data[t_labels[tid]] * 100), "details": [f"{succ}/{seen}"] + analyzer.player_song_details[name][f"Type {tid}"]} if pd.notnull(row_data[t_labels[tid]]) else np.nan
                row[delta_key]      = float(row_data[delta_key]) if pd.notnull(row_data.get(delta_key)) else np.nan

            if watched:
                succ_rig    = analyzer.p_rigs_h [name]
                tot_rig     = analyzer.p_rigs   [name]
                succ_off    = analyzer.c_counts [name] - succ_rig
                tot_off     = analyzer.s_part   [name] - tot_rig

                analyzer.player_song_details[name]["Rigs"].sort(key=lambda s: s[2:].strip().lower())

                rig_song_lines  = {s[2:]    for s in analyzer.player_song_details[name]["Rigs"]}
                off_details     = [s        for s in analyzer.player_song_details[name]["Overall"] if s[2:] not in rig_song_lines]

                row.update({
                    "Rigs"          : int   (row_data["Rigs"]),
                    "Rig Rate"      : float (row_data["Rig Rate"]           * 100),
                    "Solo Rigs"     : int   (row_data["Solo Rigs"]),
                    "Solo Rig Rate" : float (row_data["Solo Rig Rate"]      * 100),
                    "Rig Over-8"    : float (row_data["Rig Over-8"])    if pd.notnull(row_data["Rig Over-8"])   else np.nan,
                    "Over-8 Δ"      : float (row_data["Over-8 Δ"])      if pd.notnull(row_data["Over-8 Δ"])     else np.nan,
                    "Rig GR"        : {"count": float(row_data["Rig GR"]    * 100), "details": [f"{succ_rig}/{tot_rig}"] + analyzer.player_song_details[name]["Rigs"]}  if pd.notnull(row_data["Rig GR"]) else np.nan,
                    "Off GR"        : {"count": float(row_data["Off GR"]    * 100), "details": [f"{succ_off}/{tot_off}"] + off_details}                                 if pd.notnull(row_data["Off GR"]) else np.nan,
                    "Rig Δ"         : float (row_data["Rig Δ"]              * 100),
                })

            h_diffs = analyzer.p_hit_diff       .get(name, [])
            h_vints = analyzer.p_hit_vint       .get(name, [])
            times   = analyzer.p_answer_times   .get(name, [])

            if h_diffs: row["Mean Difficulty Hit"] = {
                "count"     : float(np.mean(h_diffs)),
                "details"   : [
                    f"Minimum: {float(np.min(h_diffs)):.2f}",
                    f"Median: {float(np.median(h_diffs)):.2f}",
                    f"Maximum: {float(np.max(h_diffs)):.2f}",
                    f"Standard Deviation: {float(np.std(h_diffs)):.2f}"
                ]
            }
            else: row["Mean Difficulty Hit"] = np.nan

            if h_vints: row["Median Vintage Hit"] = {
                "count"     : float(np.median(h_vints)),
                "details"   : [
                    f"Minimum: {float(np.min(h_vints))}",
                    f"Mean: {float(np.mean(h_vints)):.2f}",
                    f"Maximum: {float(np.max(h_vints))}",
                    f"Standard Deviation: {float(np.std(h_vints)):.2f}"
                ]
            }
            else: row["Median Vintage Hit"] = np.nan

            if times: row["Median Time"] = {
                "count"     : float(row_data["Median Time"]),
                "details"   : [
                    f"Minimum: {float(np.min(times)):.2f}",
                    f"Mean: {float(np.mean(times)):.2f}",
                    f"Maximum: {float(np.max(times)):.2f}",
                    f"Standard Deviation: {float(np.std(times)):.2f}",
                ]
            }
            else: row["Median Time"] = np.nan

            seen_chan       = analyzer.p_chan_s[name]
            succ_chan       = analyzer.p_chan_c[name]
            row["Chant GR"] = {"count": float(row_data["Chant GR"] * 100), "details": [f"{succ_chan}/{seen_chan}"] + analyzer.player_song_details[name]["Chant"]} if pd.notnull(row_data["Chant GR"]) else np.nan

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
            if df_players[c_field].isna().all() or (df_players[c_field].astype(str) == "nan").all():
                df_players = df_players.drop(columns=[c_field])

    if "GR" in df_players.columns and "Eru" not in analyzer.tour_label:
        if analyzer.val_str == "default":
            if      analyzer.tour_label == "Watched 2+8s"   : th_val = "25, 20, 15, 10, 5"
            elif    "Watched" in analyzer.tour_label        : th_val = "28, 18, 12, 6"
            else                                            : th_val = "28, 19, 8"
        else                                                : th_val = analyzer.val_str

        try                 : th = [float(x.strip()) for x in th_val.split(",")] if th_val else []
        except Exception    : th = [28.0, 18.0, 12.0, 6.0]

        gv = df_players["GR"].map(lambda x: x["count"] if isinstance(x, dict) else x).tolist()

        for t in th:
            f_idx = -1

            for i, v in enumerate(gv):
                if pd.notnull(v) and v >= t: f_idx = i

            if f_idx != -1 and f_idx < len(df_players) - 1: borders.append(int(f_idx))

    desc_cols = [
        "Elo", "GR", "GR Δ", "UF", "UF Δ", "Score",
        "1/8s", "2/8s", "Lives Taken", "Lives Saved",
        "OP GR", "OP Δ", "ED GR", "ED Δ", "IN GR", "IN Δ",
        "Rigs", "Rig Rate", "Solo Rigs", "Solo Rig Rate",
        "Over-8 Δ", "Rig GR", "Off GR", "Rig Δ",
        "Median Vintage Hit", "Chant GR"
    ]

    asc_cols    = ["7/8s", "Median Time", "Mean Over-8", "Rig Over-8", "Mean Difficulty Hit"]
    int_cols    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs", "Solo Rigs"]
    rate_cols   = ["GR", "OP GR", "ED GR", "IN GR", "Chant GR", "Rig GR", "Off GR"]
    stats_hl    = {}

    elo_ser     = df_players["Elo"].fillna(0.0) if "Elo" in df_players.columns else pd.Series(0.0, index = df_players.index)
    gr_ser      = df_players["GR"]      .map(lambda x: x["count"] if isinstance(x, dict) else x).fillna(0.0)
    rig_ser     = df_players["Rigs"]    .map(lambda x: x["count"] if isinstance(x, dict) else x).fillna(0.0) if "Rigs" in df_players.columns else pd.Series(0.0, index = df_players.index)
    mask_series = pd.Series(eligibility, index=df_players.index)

    for col in df_players.columns:
        if col in ["Team", "Tier"]: continue

        if col in desc_cols or col in asc_cols:
            if col in int_cols or col in rate_cols or col in ["Mean Difficulty Hit", "Median Vintage Hit", "Median Time"]   : num = df_players[col].map(lambda x: x["count"] if isinstance(x, dict) else x)
            else                                                                                                            : num = df_players[col]

            num     = pd.to_numeric(num, errors="coerce")
            el_num  = num[mask_series].dropna() if col in int_cols else num.dropna()

            if not num.dropna().empty:
                if col == "Chant GR" and num.dropna().max() == 0: continue

                if col in desc_cols:
                    best_val    = num.dropna().max()
                    worst_val   = el_num.min() if not el_num.empty else None

                else:
                    best_val = num.dropna().min()

                    if col == "Median Time":
                        under_limit = el_num[el_num < THRESH_TIME]
                        worst_val   = under_limit   .max() if not under_limit   .empty else None
                    else: worst_val = el_num        .max() if not el_num        .empty else None

                best_b_idx  = num       [num    == best_val]    .index if pd.notnull(best_val)  else pd.Index([])
                worst_b_idx = el_num    [el_num == worst_val]   .index if pd.notnull(worst_val) else pd.Index([])

                if col == "Solo Rigs":
                    best_idx    = int(rig_ser.loc[best_b_idx]   .idxmin()) if not best_b_idx    .empty else None
                    worst_idx   = int(rig_ser.loc[worst_b_idx]  .idxmax()) if not worst_b_idx   .empty else None

                elif col == "Solo Rig Rate":
                    best_idx    = int(rig_ser.loc[best_b_idx]   .idxmax()) if not best_b_idx    .empty else None
                    worst_idx   = int(rig_ser.loc[worst_b_idx]  .idxmax()) if not worst_b_idx   .empty else None

                elif col == "Elo":
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

                stats_hl[col] = {"best_idx": best_idx, "worst_idx": worst_idx}

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

def render_dashboard_tour(analyzer, watched: bool, tour_song_details: dict, player_song_details: dict) -> list[dict]:
    stats           = compute_tour_stats(analyzer, analyzer.use_teams, watched)
    tour_unrolled   = []

    for row in stats:
        metric_name = row[0]
        display_val = str(row[1])
        link_key    = row[2]
        details     = []

        if link_key is not None:
            if      isinstance(link_key, str): details = tour_song_details.get(link_key, [])
            elif    isinstance(link_key, tuple):
                stat_key, player_name   = link_key
                lookup_key              = "Solo Rig Conversions" if "Converter" in metric_name else stat_key
                details                 = player_song_details.get(player_name, {}).get(lookup_key, [])

        details.sort(key = str.lower)
        tour_unrolled.append({"Metric": metric_name, "Value": {"count": display_val, "details": details}})

    return tour_unrolled

def render_dashboard_team(analyzer, df_teams: pd.DataFrame, team_song_details: dict) -> tuple[list[dict], dict]:
    team_rows, team_hl_rules = [], {}
    if not analyzer.use_teams: return team_rows, team_hl_rules

    for _, row_data in df_teams.iterrows():
        tid = row_data["_tid"]
        h   = row_data["_history"]

        team_song_details[tid]["Total 1/8s"].sort(key = str.lower)
        details_hover = []

        if h["total_matches"] > 0:
            def line_fmt(header, sub_dict):
                if not sub_dict: return ""
                items = [f"{opp} ({', '.join(scores)})" for opp, scores in sub_dict.items()]
                return f"{header}: {', '.join(items)}"

            w_line = line_fmt("Win", h["wins"])
            l_line = line_fmt("Loss", h["losses"])
            t_line = line_fmt("Tie", h["ties"])

            if w_line: details_hover.append(w_line)
            if l_line: details_hover.append(l_line)
            if t_line: details_hover.append(t_line)

        summary_parts   = [int(x) for x in h["summary"].split("-")]
        tot_m           = sum(summary_parts)
        tie_val         = summary_parts[2] if len(summary_parts) > 2 else 0
        win_rate_val    = ((summary_parts[0] + 0.5 * tie_val) / tot_m * 100) if tot_m > 0 else np.nan

        item_payload = {
            "Team Leader"   : row_data["Team Leader"],
            "Mean Elo"      : float(row_data["Mean Elo"]),
            "Mean GR"       : float(row_data["Mean GR"]),
            "Total 1/8s"    : {"count": int(row_data["Total 1/8s"]), "details": team_song_details[tid]["Total 1/8s"]},
            "Mean Over-8"   : float(row_data["Mean Over-8"]),
            "Rig Synergy"   : float(row_data["Rig Synergy"]),
            "Off Synergy"   : float(row_data["Off Synergy"]),
            "Shared Rigs"   : float(row_data["Shared Rigs"]),
        }

        if h["total_matches"] > 0:
            item_payload["Win Record"]      = {"count": h["summary"], "details": details_hover}
            item_payload["_win_pct_sort"]   = win_rate_val
        else: item_payload["_win_pct_sort"] = -1.0

        team_rows.append(item_payload)

    if team_rows:
        df_teams_temp   = pd.DataFrame(team_rows)
        desc            = ["Mean Elo", "Mean GR", "Total 1/8s", "Rig Synergy", "Off Synergy", "Shared Rigs", "Win Record"]
        asc             = ["Mean Over-8"]

        for col in df_teams_temp.columns:
            if col.startswith("_"): continue
            num = df_teams_temp["_win_pct_sort"] if col == "Win Record" else df_teams_temp[col].map(lambda x: x["count"] if isinstance(x, dict) else x)

            if not num.dropna().empty and (col in desc or col in asc):
                clean_num = num.loc[df_teams_temp["_win_pct_sort"] != -1.0] if col == "Win Record" else num
                if clean_num.dropna().empty: continue

                best_val    = clean_num.dropna().min() if col in asc else clean_num.dropna().max()
                worst_val   = clean_num.dropna().max() if col in asc else clean_num.dropna().min()

                best_b_idx  = num[num == best_val]  .index
                worst_b_idx = num[num == worst_val] .index

                team_hl_rules[col] = {
                    "best_idx"  : int(best_b_idx    [0]) if not best_b_idx  .empty else None,
                    "worst_idx" : int(worst_b_idx   [0]) if not worst_b_idx .empty else None,
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

def render_dashboard_tier(analyzer, rows1: list[dict], rows2: list[dict], player_song_details: dict) -> dict:
    tier_data = {}

    for r1, r2 in zip(rows1, rows2):
        tr              = r1["Tier"]
        tier_data[tr]   = []
        players_tracked = {p["player"] for p in r1["_players"]["gen"]}

        for p in players_tracked:
            tot = analyzer.s_part   [p]
            cor = analyzer.c_counts [p]
            chc = analyzer.p_chan_c [p]
            cht = analyzer.p_chan_s [p]

            gen = 100 * cor / tot           if tot else 0.0
            atk = next((x["value"] for x in r1["_players"]["atk"] if x["player"] == p), 0.0)
            blk = next((x["value"] for x in r1["_players"]["blk"] if x["player"] == p), 0.0)
            con = 100 * (atk + blk) / cor   if cor else 0.0
            spd = next((x["value"] for x in r2["_players"]["spd"] if x["player"] == p), None)
            chn = 100 * chc / cht           if cht else 0.0

            player_song_details[p]["Lives Taken"].sort(key = str.lower)
            player_song_details[p]["Lives Saved"].sort(key = str.lower)

            taken_suffix, saved_suffix = {}, {}

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

            times = analyzer.p_answer_times.get(p, [])

            if times and spd is not None and pd.notnull(spd): t_det = {
                "count"     : float(round(spd, 2)),
                "details"   : [
                    f"Minimum: {float(np.min(times)):.2f}",
                    f"Mean: {float(np.mean(times)):.2f}",
                    f"Maximum: {float(np.max(times)):.2f}",
                    f"Standard Deviation: {float(np.std(times)):.2f}"
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

def render_dashboard_song(analyzer) -> list[dict]:
    song_matrix_list = []

    for s in analyzer.song_data:
        if s["vintage"] > 0: song_matrix_list.append({
            "vintage"       : float(round(s["vintage"],     2)),
            "difficulty"    : float(round(s["difficulty"],  2)),
            "correct_count" : int(s["correct_count"]),
        })

    return song_matrix_list

def render_dashboard_plot(analyzer, avg_rank: float, raw_vintage_by_guess: dict, raw_vintage_by_list: dict) -> tuple[list[dict], list[dict]]:
    pool_data = []

    for name in analyzer.s_part:
        if analyzer.c_counts[name] > 0:
            tot         = analyzer.s_part[name]
            uf_scaled   = (analyzer.p_usefulness_sum[name] * avg_rank * 8) / tot    if tot else 0.0
            gr_val      = analyzer.c_counts[name] / tot                             if tot else 0.0

            try                 : elo = float(analyzer.elo_map.get(name.lower(), 0.0))
            except Exception    : elo = 0.0

            pool_data.append({"name": name, "uf": uf_scaled, "gr": gr_val, "elo": elo})

    els = np.array([p["elo"]    for p in pool_data])
    ufs = np.array([p["uf"]     for p in pool_data])
    grs = np.array([p["gr"]     for p in pool_data])

    if len(els) > 1 and np.var(els) > 0:
        slope_uf, int_uf = np.polyfit(els, ufs, 1)
        slope_gr, int_gr = np.polyfit(els, grs, 1)

        res_std_uf = np.std(ufs - (slope_uf * els + int_uf))
        res_std_gr = np.std(grs - (slope_gr * els + int_gr))

        if res_std_uf == 0: res_std_uf = 1
        if res_std_gr == 0: res_std_gr = 1

    else:
        slope_uf, int_uf, res_std_uf = 0, np.mean(ufs) if len(ufs) > 0 else 0, 1
        slope_gr, int_gr, res_std_gr = 0, np.mean(grs) if len(grs) > 0 else 0, 1

    scatter_list, arrow_list = [], []

    for name in analyzer.s_part:
        if analyzer.c_counts[name] > 0:
            yl = np.median(analyzer.p_l_vint[name]) if analyzer.p_l_vint[name] else np.nan
            yg = np.median(analyzer.p_c_vint[name]) if analyzer.p_c_vint[name] else np.nan

            p_vints     = raw_vintage_by_guess.get(name, [])
            p_vint_med  = np.median([extract_year(v) for v in p_vints]) if p_vints else yg
            p_seas      = format_year(p_vint_med)                       if p_vints else f"Winter {int(yg)}" if pd.notnull(yg) else "N/A"

            r_vints     = raw_vintage_by_list.get(name, [])
            r_vint_med  = np.median([extract_year(v) for v in r_vints]) if r_vints else yl
            r_seas      = format_year(r_vint_med)                       if r_vints else f"Winter {int(yl)}" if pd.notnull(yl) else "N/A"

            tot         = analyzer.s_part[name]
            uf_scaled   = (analyzer.p_usefulness_sum[name] * avg_rank * 8) / tot    if tot else 0.0
            gr_val      = analyzer.c_counts[name] / tot                             if tot else 0.0

            try                 : elo = float(analyzer.elo_map.get(name.lower(), 0.0))
            except Exception    : elo = 0.0

            residual_uf     = uf_scaled - (slope_uf * elo + int_uf)
            residual_gr     = gr_val - (slope_gr * elo + int_gr)
            perf_score_uf   = (1 / (1 + np.exp(SCALE_PERF * (residual_uf / res_std_uf)))) * 100
            perf_score_gr   = (1 / (1 + np.exp(SCALE_PERF * (residual_gr / res_std_gr)))) * 100
            perf_score      = (perf_score_uf * 0.5) + (perf_score_gr * 0.5)

            base_node = {
                "acronym"           : analyzer.player_acronyms.get(name.lower(), name[:3].upper()),
                "name"              : name,
                "over8"             : float(round(analyzer.p_overs_sum[name]    / analyzer.c_counts [name], 2)),
                "vintage"           : float(round(p_vint_med, 2)),
                "seasonal_vintage"  : p_seas,
                "gr"                : float(round(analyzer.c_counts[name]       / analyzer.s_part   [name] * 100, 2)) if analyzer.s_part[name] else 0.0,
                "rig_gr"            : float(round(analyzer.p_rigs_h[name]       / analyzer.p_rigs   [name] * 100, 2)) if analyzer.p_rigs[name] else 0.0,
                "performance"       : float(round(perf_score, 2)),
                "rig_rate"          : float(round(analyzer.p_rigs[name]         / analyzer.s_part   [name] * 100, 2)) if analyzer.s_part[name] else 0.0
            }

            scatter_list.append(base_node)

            if analyzer.p_l_corr[name] and pd.notnull(yl) and pd.notnull(yg):
                hit_over8   = np.mean(analyzer.p_lh_corr[name]) if analyzer.p_lh_corr[name] else base_node["over8"]
                hit_vint    = np.median(analyzer.p_lh_vint[name]) if analyzer.p_lh_vint[name] else base_node["vintage"]

                arrow_list.append({
                    "acronym"                   : base_node["acronym"],
                    "name"                      : name,
                    "x_start"                   : float(round(np.mean(analyzer.p_l_corr[name]), 2)),
                    "y_start"                   : float(round(r_vint_med,                       2)),
                    "seasonal_vintage_start"    : r_seas,
                    "x_end"                     : base_node["over8"],
                    "y_end"                     : base_node["vintage"],
                    "x_hit"                     : float(round(hit_over8,                        2)),
                    "y_hit"                     : float(round(hit_vint,                         2)),
                    "seasonal_vintage_end"      : p_seas,
                    "rig_gr"                    : base_node["rig_gr"],
                    "gr"                        : base_node["gr"],
                    "rig_rate"                  : base_node["rig_rate"],
                })

    return scatter_list, arrow_list

def create_dashboard_html(analyzer, path: Path, use_teams: bool, watched: bool):
    active = [t for t in [1, 2, 3] if any(analyzer.p_type_s[p][t] > 0 for p in analyzer.s_part)]
    if len(active) <= 1: active = []

    t_labels    = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
    valid_elos  = [float(v) for v in analyzer.elo_map.values() if str(v).replace(".", "", 1).isdigit() or (str(v).startswith("-") and str(v)[1:].replace(".", "", 1).isdigit())]
    avg_rank    = np.mean(valid_elos)   if valid_elos           else 1.0
    team_count  = len(analyzer.rosters) if analyzer.use_teams   else 0

    if team_count <= 2  : stage = "Final" if analyzer.base_exp >= 3 else f"R{analyzer.base_exp}"
    else                : stage = "Mid-Tour" if analyzer.base_exp == 3 else "Final" if (team_count <= 4 and analyzer.base_exp >= 6) or (team_count > 4 and analyzer.base_exp >= 5) else f"R{analyzer.base_exp}"

    prefix      = f"{analyzer.tour_label.strip()} Tour: {stage}"
    diffs       = [s["difficulty"] for s in analyzer.song_data]
    max_diff    = max(diffs)    if diffs            else 0
    num_x       = 8             if max_diff < 40    else 9
    num_y       = 8             if max_diff < 40    else 9
    df_base, _  = compute_player_rows(analyzer, analyzer.elo_map, analyzer.apps, analyzer.exp_map, analyzer.base_exp, analyzer.new_players, watched, active, t_labels, avg_rank)

    def player_sort_key(x):
        gr = (analyzer.c_counts[x] / analyzer.s_part[x]) if analyzer.s_part[x] else 0.0

        try                 : elo = float(analyzer.elo_map.get(x.lower(), float("inf")))
        except Exception    : elo = float("inf")

        return (gr, -elo)

    sorted_players  = sorted            (analyzer.s_part.keys(), key = player_sort_key, reverse = True)
    df_teams        = compute_team_rows (analyzer, analyzer.assignments, analyzer.t1_lookup)
    rows1, rows2    = compute_tier_rows (analyzer, analyzer.assignments, any(analyzer.p_chan_s.values()))

    render_players, render_hl_rules, render_borders, render_eligibility = render_dashboard_player(
        analyzer,
        sorted_players,
        active,
        t_labels,
        watched,
        analyzer.player_song_details,
        df_base,
    )

    render_teams,       render_team_hl_rules    = render_dashboard_team(analyzer, df_teams, analyzer.team_song_details)
    render_scatter,     render_arrows           = render_dashboard_plot(analyzer, avg_rank, analyzer.raw_vintage_by_guess, analyzer.raw_vintage_by_list)
    render_tour_stats                           = render_dashboard_tour(analyzer, watched, analyzer.tour_song_details, analyzer.player_song_details)
    render_tier_merged                          = render_dashboard_tier(analyzer, rows1, rows2, analyzer.player_song_details)
    render_songs                                = render_dashboard_song(analyzer)

    explanations = {
        "Player"                    : "★ New player<br>▲ Subbed in<br>▼ Subbed out",
        "UF"                        : "Usefulness<br>Calculates this player's contribution to their team, scaled by Elo and songs played",
        "Score"                     : "Calculates this player's Elo and GR against what's expected from their Elo<br>50 means this player is playing to expectations",
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
        "Worst Solo Rig Converter"  : "100 * Solo from Solo Rig / Solo Rig<br>Shows the worst player at converting their own solo rig into a solo",
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
        "json_matrix_songs"     : analyzer.matrix_song_details,
        "json_scatter"          : render_scatter,
        "json_arrows"           : render_arrows,
        "json_explanations"     : explanations,
        "generated_timestamp"   : int(time.time() * 1000),
    }

    search_songs_list = []

    for path_json in analyzer.json_paths:
        try:
            with open(path_json, encoding = "utf-8") as f: data_j = json.load(f)
        except Exception: continue

        raw_f_players = set()

        for s in data_j.get("songs", []):
            for p in s.get("correctGuessPlayers", []):
                if      isinstance(p, str)                  : raw_f_players.add(p)
                elif    isinstance(p, dict) and "name" in p : raw_f_players.add(p["name"])

            for ls in s.get("listStates", []):
                if "name" in ls: raw_f_players.add(ls["name"])

        final_members = set(raw_f_players)

        if analyzer.use_teams:
            t_in_f = {analyzer.assignments[p.lower()][0] for p in raw_f_players if p.lower() in analyzer.assignments}

            for tid in t_in_f:
                for m_p in analyzer.rosters[tid]:
                    if m_p.lower() not in analyzer.assignments:
                        for c_p in raw_f_players:
                            if (c_p.lower() in analyzer.assignments and analyzer.assignments[c_p.lower()][0] == tid):
                                analyzer.assignments[m_p.lower()] = analyzer.assignments[c_p.lower()]

            if len(final_members) < 8:
                for tid in t_in_f: final_members.update(analyzer.rosters[tid])

        room_players_list = sorted(list(final_members))

        for song in data_j.get("songs", []):
            si          = song.get("songInfo", {})
            ann_id_raw  = si.get("annId")

            if not ann_id_raw: continue

            lives_taken_players, lives_saved_players    = [], []
            anime_romaji                                = si.get("animeNames",      {}).get("romaji",   "Unknown")  .strip()
            anime_english                               = si.get("animeNames",      {}).get("english",  "")         .strip()
            song_name                                   = si.get("songName",        "Unknown")                      .strip()
            raw_artist                                  = si.get("artist",          "Unknown")                      .strip()
            artist_arr                                  = [a.strip() for a in raw_artist.split(",") if a            .strip()]   if raw_artist               else []
            composer_name                               = si.get("composerInfo",    {}).get("name",     "Unknown")  .strip()    if si.get("composerInfo")   else "Unknown"
            arranger_name                               = si.get("arrangerInfo",    {}).get("name",     "Unknown")  .strip()    if si.get("arrangerInfo")   else "Unknown"

            st          = si.get("type",        3)
            t_num       = si.get("typeNumber",  0)
            type_fmt    = f"Opening {t_num}" if st == 1 else f"Ending {t_num}" if st == 2 else "Insert"

            ann_url         = f"https://www.animenewsnetwork.com/encyclopedia/anime.php?id={str(ann_id_raw)}"
            anime_type_raw  = str(si.get("animeType", "N/A")).strip()
            anime_type      = "Movie" if anime_type_raw.lower() == "movie" else "Special" if anime_type_raw.lower() == "special" else anime_type_raw
            vint_raw        = str(si.get("vintage", "Unknown")).strip().replace("\n", " ").replace("\r", " ")

            try:
                diff_val    = si.get("animeDifficulty")
                safe_diff   = f"{float(diff_val):.2f}" if diff_val is not None and float(diff_val) > 0 else "Unrated"
            except Exception: safe_diff = "Unrated"

            video_url       = song.get("videoUrl",              "")
            raw_correct     = song.get("correctGuessPlayers",   [])
            ann_song_id_str = str(si.get("annSongId",           ""))
            is_chanting_str = "Yes" if ann_song_id_str in analyzer.chanting_ids else "No"
            guess_times     = {}

            for p in raw_correct:
                if isinstance(p, dict) and "name" in p:
                    t_val = p.get("answerTime")
                    guess_times[p["name"].lower()] = f"{float(t_val):.2f}" if t_val is not None else "N/A"
                elif isinstance(p, str): guess_times[p.lower()] = "N/A"

            guessers_flat   = [p if isinstance(p, str) else p["name"] for p in raw_correct if isinstance(p, (str, dict))]
            raw_lists       = song.get("listStates", [])
            listers_flat    = [ls["name"] for ls in raw_lists if isinstance(ls, dict) and "name" in ls]

            if analyzer.use_teams:
                t_list = list({analyzer.assignments[p.lower()][0] for p in raw_f_players if p.lower() in analyzer.assignments})

                if len(t_list) == 2:
                    tA, tB      = t_list[0], t_list[1]
                    correct_set = set(guessers_flat)

                    for cur, opp in [(tA, tB), (tB, tA)]:
                        cC, oC = correct_set & analyzer.rosters[cur], correct_set & analyzer.rosters[opp]

                        if not oC:
                            for p in cC: lives_taken_players.append(p.lower())

                        if len(cC) == 1 and len(oC) > 0:
                            lone_p = list(cC)[0]
                            lives_saved_players.append(lone_p.lower())

            def group_by_team_structure(target_players, include_times = False):
                if not analyzer.use_teams:
                    sorted_p = sorted(target_players, key = str.lower)
                    if include_times: return [f"{p} ({guess_times.get(p.lower(), 'N/A')})" for p in sorted_p]
                    return sorted_p

                team_buckets = defaultdict(list)

                for p in target_players:
                    tid, _ = analyzer.assignments.get(p.lower(), (None, "5"))
                    team_buckets[tid].append(p)

                sorted_tids = sorted([t for t in team_buckets.keys() if t is not None])
                hover_lines = []

                for tid in sorted_tids:
                    leader_name = analyzer.t1_lookup.get(tid, f"Team {tid}")
                    pts_sorted  = sorted(team_buckets[tid], key = lambda x: analyzer.assignments.get(x.lower(), (None, "5"))[1])
                    p_strings   = [f"{p} ({guess_times.get(p.lower(), 'N/A')})" for p in pts_sorted] if include_times else pts_sorted

                    hover_lines.append(f"Team {leader_name}: {', '.join(p_strings)}")

                all_active_tids = {analyzer.assignments[p.lower()][0] for p in final_members if p.lower() in analyzer.assignments}
                missing_tids    = sorted(list(all_active_tids - set(team_buckets.keys())))

                for tid in missing_tids:
                    leader_name = analyzer.t1_lookup.get(tid, f"Team {tid}")
                    hover_lines.append(f"Team {leader_name}: None")

                return hover_lines

            guessers_hover  = group_by_team_structure(guessers_flat,    include_times = True)
            listers_hover   = group_by_team_structure(listers_flat,     include_times = False)

            raw_genres = si.get("animeGenre", []) if isinstance(si.get("animeGenre"), list) else []
            raw_tags = (
                [t for t in si.get("animeTags", []) if t not in EXCLUDED_TAGS]
                if isinstance(si.get("animeTags"), list)
                else []
            )
            raw_alts        = si.get("altAnimeNames", []) if isinstance(si.get("altAnimeNames"), list) else []
            raw_alts_ans    = si.get("altAnimeNamesAnswers", []) if isinstance(si.get("altAnimeNamesAnswers"), list) else []
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
                "correct_teams_flat"    : [analyzer.assignments[p.lower()][0] for p in guessers_flat if p.lower() in analyzer.assignments],
                "alts"                  : combined_alts,
                "start_sample"          : song.get("startPoint", 0)
            })

    search_songs_list.sort(key=lambda x: x["romaji"].lower())
    (path / "jsons").mkdir(parents = True, exist_ok = True)

    with open(path / "jsons" / "search.json",   "w", encoding = "utf-8") as f: json.dump(search_songs_list, f, ensure_ascii = False, indent = 4)
    with open(path / "jsons" / "data.json",     "w", encoding = "utf-8") as f: json.dump(data_payload,      f, ensure_ascii = False, indent = 4)

    template_dir = analyzer.script_dir / "help" / "template"

    shutil.copy     (template_dir / "index.html",           path / "index.html")
    shutil.copy     (template_dir / "styles.css",           path / "styles.css")
    shutil.copy     (template_dir / "config.js",            path / "config.js")
    shutil.copy     (template_dir / "jsons" / "name.json",  path / "jsons" / "name.json")
    shutil.copytree (template_dir / "tabs",                 path / "tabs", dirs_exist_ok = True)