import math, matplotlib, os
import matplotlib.colors    as mc
import matplotlib.pyplot    as plt
import numpy                as np
import pandas               as pd

matplotlib.use('Agg')

from .config        import *
from .statistics    import *
from adjustText     import adjust_text
from html2image     import Html2Image
from pathlib        import Path
from PIL            import Image
from scipy.spatial  import ConvexHull

def create_player_png(
    analyzer,
    elo_map     : dict,
    watched     : bool,
    stage       : str,
    path        : Path,
    apps        : dict,
    prefix      : str,
    exp_map     : dict,
    base_exp    : int,
    new_players : list,
    val_str     : str,
):
    t_labels    = {1: "OP GR", 2: "ED GR", 3: "IN GR"}
    active      = [t for t in [1, 2, 3] if any(analyzer.p_type_s[p][t] > 0 for p in analyzer.s_part)]

    if len(active) <= 1: active = []

    valid_elos  = [float(v) for v in elo_map.values() if str(v).replace(".", "", 1).isdigit() or (str(v).startswith("-") and str(v)[1:].replace(".", "", 1).isdigit())]
    avg_rank    = np.mean(valid_elos) if valid_elos else 1.0
    df, mask    = compute_player_rows(analyzer, elo_map, apps, exp_map, base_exp, new_players, watched, active, t_labels, avg_rank)
    df_png      = df.copy()
    pcts        = (["GR"] + [t_labels[t] for t in active] + (["Rig Rate", "Solo Rig Rate", "Rig Δ", "Rig GR", "Off GR"] if watched else []) + ["Chant GR"])

    if "Elo"            in df_png.columns: df_png["Elo"]            = pd.to_numeric(df_png["Elo"],          errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    if "UF"             in df_png.columns: df_png["UF"]             = pd.to_numeric(df_png["UF"],           errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    if "Score"          in df_png.columns: df_png["Score"]          = pd.to_numeric(df_png["Score"],        errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    if "Median Time"    in df_png.columns: df_png["Median Time"]    = pd.to_numeric(df_png["Median Time"],  errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    if "Mean Over-8"    in df_png.columns: df_png["Mean Over-8"]    = pd.to_numeric(df_png["Mean Over-8"],  errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    if "Rig Over-8"     in df_png.columns: df_png["Rig Over-8"]     = pd.to_numeric(df_png["Rig Over-8"],   errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    if "Over-8 Δ"       in df_png.columns: df_png["Over-8 Δ"]       = pd.to_numeric(df_png["Over-8 Δ"],     errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

    for c in pcts: df_png[c] = (pd.to_numeric(df_png[c], errors = "coerce").mul(100).map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A"))
    delta_cols = ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]

    for dc in delta_cols:
        if dc in df_png.columns: df_png[dc] = pd.to_numeric(df_png[dc], errors="coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")

    export_png(analyzer, df_png, path, "Player.png", f"{prefix}{stage}", mask, val_str)

def create_tour_png(analyzer, use_teams: bool, watched: bool, path: Path):
    stats   = compute_tour_stats(analyzer, use_teams, watched)
    half    = (len(stats) + 1) // 2
    left    = stats[:half]
    right   = stats[half:]

    while len(right) < len(left): right.append(["", "", None])

    split_stats = [[l[0], l[1], r[0], r[1]] for l, r in zip(left, right)]
    df_tour     = pd.DataFrame(split_stats, columns = ["Metric", "Value", "Metric", "Value"])

    export_png(analyzer, df_tour, path, "Tour.png", "Tour Statistics")

def create_team_png(analyzer, assigns: dict, t1_lookup: dict, path: Path):
    analyzer.text_var_wlt   = "Yes"
    df                      = compute_team_rows(analyzer, assigns, t1_lookup)

    if "_win_pct" in df.columns : df_png = df.drop(columns=["_tid", "_history", "_win_pct"])
    else                        : df_png = df.drop(columns=["_tid", "_history"])

    if "Win Record" in df_png.columns and ((df_png["Win Record"].astype(str) == "0-0-0").all() or (df_png["Win Record"].astype(str) == "0-0").all()): df_png = df_png.drop(columns=["Win Record"])
    watched_valid = analyzer.missing_list_count <= 5

    if not watched_valid:
        df_png      = df_png.drop(columns=["Rig Synergy", "Off Synergy", "Shared Rigs"], errors="ignore")
        num_cols    = ["Mean Elo", "Mean GR", "Mean Over-8"]
    else: num_cols  = ["Mean Elo", "Mean GR", "Mean Over-8", "Rig Synergy", "Off Synergy", "Shared Rigs"]

    for c in num_cols: df_png[c] = pd.to_numeric(df_png[c], errors = "coerce").map(lambda x: f"{x:.2f}" if pd.notnull(x) else "N/A")
    export_png(analyzer, df_png, path, "Team.png", "Team Statistics")
    analyzer.text_var_wlt = "No"

def create_tier_png(analyzer, assigns: dict, path: Path, has_chanting_songs: bool):
    rows1, rows2    = compute_tier_rows(analyzer, assigns, has_chanting_songs)
    valid_spd_rows  = [i for i, r in enumerate(rows2) if pd.notnull(r["spd_val"])]
    valid_chn_rows  = [i for i, r in enumerate(rows2) if r["chn_val"] is not None and r["chn_val"] > 0]

    best_gen_idx = (len(rows1) - 1 - max(range(len(rows1)), key = lambda i: rows1[::-1][i]["gen_val"]) if rows1 else None)
    best_atk_idx = (len(rows1) - 1 - max(range(len(rows1)), key = lambda i: rows1[::-1][i]["atk_val"]) if rows1 else None)
    best_blk_idx = (len(rows1) - 1 - max(range(len(rows1)), key = lambda i: rows1[::-1][i]["blk_val"]) if rows1 else None)
    best_con_idx = (len(rows2) - 1 - max(range(len(rows2)), key = lambda i: rows2[::-1][i]["con_val"]) if rows2 else None)

    best_spd_idx = min(valid_spd_rows, key = lambda i: rows2[i]["spd_val"]) if valid_spd_rows else None
    best_chn_idx = max(valid_chn_rows, key = lambda i: rows2[i]["chn_val"]) if valid_chn_rows else None

    html_parts  = ["<tr><th>Tier</th><th>Guess Rate</th><th>Lives Taken</th><th>Lives Saved</th></tr>"]
    style_hl    = f" style='background-color: {COLOR_2}; color: white; font-weight: bold;'"

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

    html_table_content  = "".join(html_parts)
    full                = f"""<html>
        <head>
            <style>
                body                {{font-family: 'Segoe UI'; background: white; display: inline-block; margin: 0}}
                h2                  {{margin: 0 0 10px 0; font-size: 40px; text-align: center}}
                table               {{border-collapse: collapse; width: auto; border: 3px solid black}}
                th                  {{font-weight: bold; font-size: 25px; text-align: center; padding: 5px 10px; border: 1px solid black; border-bottom: 3px solid black; background-color: #f0f0f0}}
                td                  {{font-size: 25px; text-align: center; padding: 5px 10px; border: 1px solid black}}
                tr:nth-child(even)  {{background-color: #f0f0f0}}
            </style>
        </head>
        <body>
            <h2>Tier Statistics</h2>
            <table>{html_table_content}</table>
        </body>
    </html>"""

    if not analyzer.browser_path: return

    hti = Html2Image(
        size                = (2000, 2000),
        browser_executable  = analyzer.browser_path,
        output_path         = str(path),
        custom_flags        = ["--log-level=3", "--silent"],
    )

    hti.screenshot(html_str = full, save_as = "Tier.png")

    try                 : trim_whitespace(path / "Tier.png")
    except Exception    : pass

def create_scatter_png(analyzer, path: Path, list_mode: bool = False, elo_map: dict = None):
    configs = []

    if list_mode:
        plist_l = [n for n in analyzer.s_part if analyzer.p_l_corr[n]]

        if plist_l:
            x_vals_l    = [np.mean(analyzer.p_l_corr[name]) for name in plist_l]
            y_vals_l    = [np.median(analyzer.p_l_vint[name]) if analyzer.p_l_vint[name] else np.nan for name in plist_l]
            valid_l     = [(p, x, y) for p, x, y in zip(plist_l, x_vals_l, y_vals_l) if not np.isnan(y)]

            if valid_l:
                plist_l, x_vals_l, y_vals_l = zip(*valid_l)
                plist_l, x_vals_l, y_vals_l = list(plist_l), list(x_vals_l), list(y_vals_l)

                rig_rates   = [analyzer.p_rigs      [name] / analyzer.s_part[name] if analyzer.s_part[name] else 0 for name in plist_l]
                grid_grs    = [analyzer.p_rigs_h    [name] / analyzer.p_rigs[name] if analyzer.p_rigs[name] else 0 for name in plist_l]

                team_count = len(analyzer.rosters) if analyzer.use_teams else 0

                scale_l = 1.00 if team_count <= THRESH_TEAM else (0.75 if team_count <= THRESH_TEAM + 2 else 0.50)
                sizes_l = [(rate * scale_l) ** 2 * 10000 for rate in rig_rates]
                cmap_l  = mc.LinearSegmentedColormap.from_list("rig_gr_cmap", [(0.0, COLOR_0), (0.7, COLOR_0), (0.8, COLOR_1), (0.9, COLOR_2), (1.0, COLOR_2)])

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
                    "cbar_ticklabels"   : ["0", "70", "80", "90", "100"],
                    "labelpad"          : -35,
                })

    plist_g = [n for n in analyzer.s_part if analyzer.c_counts[n] > 0]

    if plist_g:
        x_vals_g    = [analyzer.p_overs_sum[name] / analyzer.c_counts[name] for name in plist_g]
        y_vals_g    = [np.median(analyzer.p_c_vint[name]) if analyzer.p_c_vint[name] else np.nan for name in plist_g]
        valid_g     = [(p, x, y) for p, x, y in zip(plist_g, x_vals_g, y_vals_g) if not np.isnan(y)]

        if valid_g:
            plist_g, x_vals_g, y_vals_g = zip(*valid_g)
            plist_g, x_vals_g, y_vals_g = list(plist_g), list(x_vals_g), list(y_vals_g)

            gr_vals = [analyzer.c_counts[name] / analyzer.s_part[name] if analyzer.s_part[name] else 0 for name in plist_g]
            if elo_map is None: elo_map = {}

            valid_elos  = [float(v) for v in elo_map.values() if str(v).replace(".", "", 1).isdigit() or (str(v).startswith("-") and str(v)[1:].replace(".", "", 1).isdigit())]
            avg_rank    = np.mean(valid_elos) if valid_elos else 1.0
            norm_perf   = compute_player_performance_scores(plist_g, analyzer, elo_map, avg_rank)
            team_count  = len(analyzer.rosters) if analyzer.use_teams else 0

            scale_g = 1.00 if team_count <= THRESH_TEAM else (0.75 if team_count <= THRESH_TEAM + 2 else 0.50)
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
                    "cbar_ticklabels"   : ["0", "100"],
                    "labelpad"          : -37.5,
            })

    if not configs: return
    all_x, all_y = [], []

    for cfg in configs:
        all_x.extend(cfg["x_vals"])
        all_y.extend(cfg["y_vals"])

    if not all_x or not all_y: return

    x_min = math.floor((min (all_x) - 0.5) * 2) / 2
    x_max = math.ceil((max  (all_x) + 0.5) * 2) / 2
    y_min = math.floor(min  (all_y) - 1.0)
    y_max = math.ceil(max   (all_y) + 1.0)

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

        sc = ax.scatter(
            cfg["x_vals"], cfg["y_vals"],
            s           = cfg["sizes"],
            c           = cfg["colors"],
            cmap        = cfg["cmap"],
            vmin        = cfg["vmin"],
            vmax        = cfg["vmax"],
            edgecolors  = "black",
            alpha       = 0.95,
            zorder      = 3,
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

                ax.plot(
                    hull_points[:, 0], hull_points[:, 1],
                    color       = "black",
                    zorder      = 1,
                    linewidth   = 0.5,
                    linestyle   = "-",
                )
            except Exception: pass

        ax.set_xlim(x_min, x_max)
        ax.set_ylim(y_min, y_max)

        x_mnc, x_mxf, x_stp = math.ceil(x_min), math.floor(x_max), 1

        ax.set_xticks(range(x_mnc,      x_mxf + x_stp, x_stp))
        ax.set_yticks(range(y_min + 1,  y_max + y_stp, y_stp))

        texts = []

        for name, x, y in zip(cfg["plist"], cfg["x_vals"], cfg["y_vals"]):
            label = analyzer.player_acronyms.get(name.lower(), name[:3].upper())
            if not label: continue

            ha_align = "left"   if x >= x_center else "right"
            va_align = "bottom" if y >= y_center else "top"

            texts.append(
                ax.text(
                    x, y, label,
                    size        = 20,
                    weight      = "bold",
                    fontname    = "Segoe UI",
                    ha          = ha_align,
                    va          = va_align,
                )
            )

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
                arrowprops              = dict(arrowstyle = "-", color = "black", shrinkA = 15)
            )

        ax.set_title    (cfg["title"],  weight = "bold", fontname = "Segoe UI", fontsize = 50, pad      = 15)
        ax.set_xlabel   ("Over-8",      weight = "bold", fontname = "Segoe UI", fontsize = 25, labelpad = 5)
        ax.set_ylabel   ("Vintage",     weight = "bold", fontname = "Segoe UI", fontsize = 25, labelpad = 5)

        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, _: str(int(val))))
        plt.setp(ax.get_yticklabels(), rotation = 90, va = "center")

        ax.tick_params(axis = "x", which = "both", length = 0, labelsize = 20, pad = 5)
        ax.tick_params(axis = "y", which = "both", length = 0, labelsize = 20, pad = 2.5)

        cbar = fig.colorbar(sc, ax = ax, pad = 0.005, aspect = 40, ticks = cfg["cbar_ticks"])
        cbar.set_label(cfg["cbar_label"], weight = "bold", fontname = "Segoe UI", fontsize = 25, labelpad = cfg["labelpad"])
        cbar.ax.set_yticklabels(cfg["cbar_ticklabels"])
        cbar.ax.tick_params(labelsize = 20, length = 0)

        ax.text(0.01, 0.99, "New\nHard", transform = ax.transAxes, color = "black", fontsize = 15, va = "top",      ha = "left",    weight = "bold", alpha = 0.75)
        ax.text(0.99, 0.99, "New\nEasy", transform = ax.transAxes, color = "black", fontsize = 15, va = "top",      ha = "right",   weight = "bold", alpha = 0.75)
        ax.text(0.01, 0.01, "Old\nHard", transform = ax.transAxes, color = "black", fontsize = 15, va = "bottom",   ha = "left",    weight = "bold", alpha = 0.75)
        ax.text(0.99, 0.01, "Old\nEasy", transform = ax.transAxes, color = "black", fontsize = 15, va = "bottom",   ha = "right",   weight = "bold", alpha = 0.75)

        ax.grid(False)
        plt.tight_layout()
        plt.savefig(path / cfg["filename"], dpi = 500)
        plt.close(fig)

        try                 : trim_whitespace(path / cfg["filename"])
        except Exception    : pass

def create_song_png(analyzer, path: Path):
    diffs       = [s["difficulty"] for s in analyzer.song_data]
    max_diff    = max(diffs) if diffs else 0

    if max_diff < 40    : num_x, num_y, font_sz = 4, 4, 90
    else                : num_x, num_y, font_sz = 5, 5, 75

    counts      = np.zeros((num_y, num_x), dtype = int)
    over8_sums  = np.zeros((num_y, num_x), dtype = float)

    for s in analyzer.song_data:
        vint = int(s["vintage"])
        if vint == 0: continue

        diff    = s["difficulty"]
        x_idx   = min(int(math.floor(diff / 10)), num_x - 1)

        if num_y == 4   : y_idx = 0 if vint < 2000 else min(int(math.floor((vint - 2000) / 10)) + 1, 3)
        else            : y_idx = min(max(int(math.floor((vint - 1980) / 10)), 0), 4)

        counts      [y_idx, x_idx] += 1
        over8_sums  [y_idx, x_idx] += s["correct_count"]

    fig, ax     = plt.subplots(figsize = (10, 10))
    cmap_song   = mc.LinearSegmentedColormap.from_list("song_cmap", [(0, COLOR_0), (0.375, COLOR_1), (0.625, COLOR_2), (1, COLOR_2)])

    for y_idx in range(num_y):
        for x_idx in range(num_x):
            count = counts[y_idx, x_idx]

            if count == 0: facecolor = "white"
            else:
                avg_over8 = over8_sums[y_idx, x_idx] / count
                facecolor = cmap_song(avg_over8 / 8.0)

            rect = plt.Rectangle((x_idx, y_idx), 1, 1, facecolor = facecolor, edgecolor = "none")
            ax.add_patch(rect)

            if count > 0:
                ax.text(
                    x_idx + 0.5, y_idx + 0.45, str(count),
                    ha          = "center",
                    va          = "center",
                    color       = "white",
                    weight      = "bold",
                    fontsize    = font_sz,
                    fontname    = "Segoe UI",
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

    ax.set_xticklabels(p_labels, fontname = "Segoe UI", fontsize= 20)
    ax.set_yticklabels(y_labels, fontname = "Segoe UI", fontsize= 20, rotation = 90, va = "center")

    ax.set_title    ("Song Statistics", weight = "bold", fontname = "Segoe UI", fontsize = 50, pad      = 15)
    ax.set_xlabel   ("Difficulty",      weight = "bold", fontname = "Segoe UI", fontsize = 25, labelpad = 5)
    ax.set_ylabel   ("Vintage",         weight = "bold", fontname = "Segoe UI", fontsize = 25, labelpad = 5)

    ax.tick_params(axis = "x", which = "both", length = 0, pad = 5)
    ax.tick_params(axis = "y", which = "both", length = 0, pad = 2.5)

    ax.grid(False)

    norm    = mc.Normalize(vmin = 0.0, vmax = 8.0)
    sm      = plt.cm.ScalarMappable(cmap = cmap_song, norm = norm)

    sm.set_array([])

    cbar = fig.colorbar(sm, ax = ax, pad = 0.005, aspect = 40, ticks = [0, 3, 5, 8])
    cbar.set_label("Over-8", weight = "bold", fontname = "Segoe UI", fontsize = 25, labelpad = -12.5)
    cbar.ax.set_yticklabels(["0", "3", "5", "8"])
    cbar.ax.tick_params(labelsize = 20, length = 0)

    plt.tight_layout()
    plt.savefig(path / "Song.png", dpi = 500)
    plt.close(fig)

    try                 : trim_whitespace(path / "Song.png")
    except Exception    : pass

def export_png(analyzer, df: pd.DataFrame, path: Path, fname: str, title: str, mask: list = None, val_str: str = "default"):
    if not analyzer.browser_path: return
    df = df.reset_index(drop = True)

    delta_check_cols    = ["GR Δ", "UF Δ", "OP Δ", "ED Δ", "IN Δ"]
    cols_to_drop        = [c for c in delta_check_cols if c in df.columns and (df[c].isna() | (df[c].astype(str).str.strip() == "N/A")).all()]

    if cols_to_drop: df = df.drop(columns = cols_to_drop)

    desc = [
        "Elo", "GR", "GR Δ", "UF", "UF Δ", "Score",
        "1/8s", "2/8s", "Lives Taken", "Lives Saved",
        "OP GR", "OP Δ", "ED GR", "ED Δ", "IN GR", "IN Δ",
        "Rigs", "Rig Rate", "Solo Rigs", "Solo Rig Rate",
        "Over-8 Δ", "Rig GR", "Off GR", "Rig Δ",
        "Median Vintage Hit", "Chant GR",
        "Mean Elo", "Mean GR", "Total 1/8s", "Rig Synergy",
        "Off Synergy", "Shared Rigs", "Win Record"
    ]

    asc     = ["7/8s", "Median Time", "Mean Over-8", "Rig Over-8", "Mean Difficulty Hit"]
    rest    = ["1/8s", "2/8s", "7/8s", "Lives Taken", "Lives Saved", "Rigs"]
    stats   = {}
    elo_col = "Elo" if "Elo" in df.columns else "Mean Elo" if "Mean Elo" in df.columns else None
    elo_ser = pd.to_numeric(df[elo_col], errors = "coerce").fillna(0.0) if elo_col else pd.Series(0.0, index = df.index)

    for col in df.columns:
        if col in desc or col in asc:
            if col == "Win Record":
                def parse_wlt(val):
                    try:
                        parts   = [int(x) for x in str(val).split("-")]
                        total   = sum(parts)
                        ties    = parts[2] if len(parts) > 2 else 0

                        return ((parts[0] + 0.5 * ties) / total) if total > 0 else -1.0
                    except Exception: return -1.0

                num     = df[col].apply(parse_wlt)
            else: num   = pd.to_numeric(df[col].astype(str).str.replace("%", ""), errors = "coerce")

            el_num = num[mask].dropna() if mask is not None and col in rest else num.dropna()

            if not num.dropna().empty:
                if col == "Chant GR" and num.dropna().max() == 0: continue

                if col in desc:
                    best_val    = num.dropna().max()
                    worst_val   = el_num.min() if not el_num.empty else None

                else:
                    best_val = num.dropna().min()

                    if col == "Median Time":
                        under_limit = el_num[el_num < THRESH_TIME]
                        worst_val   = under_limit.max() if not under_limit  .empty else None
                    else: worst_val = el_num.max()      if not el_num       .empty else None

                best_b_indices  = num       [num    == best_val]    .index if pd.notnull(best_val)  else pd.Index([])
                worst_b_indices = el_num    [el_num == worst_val]   .index if pd.notnull(worst_val) else pd.Index([])

                el_cols = ["Elo", "Mean Elo"]
                gr_cols = ["OP GR", "OP Δ", "ED GR", "ED Δ", "IN GR", "IN Δ", "Chant GR"]
                rig_ser = pd.to_numeric(df["Rigs"], errors = "coerce").fillna(0) if "Rigs" in df.columns else pd.Series(0, index = df.index)

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
                    best_idx    = pd.to_numeric(df["GR"], errors = "coerce").fillna(0).loc[best_b_indices]  .idxmin() if not best_b_indices     .empty else None
                    worst_idx   = pd.to_numeric(df["GR"], errors = "coerce").fillna(0).loc[worst_b_indices] .idxmax() if not worst_b_indices    .empty else None

                elif col == "Rig GR" and "Rigs" in df.columns:
                    best_idx    = rig_ser.loc[best_b_indices]   .idxmax() if not best_b_indices     .empty else None
                    worst_idx   = elo_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                else:
                    best_idx    = elo_ser.loc[best_b_indices]   .idxmin() if not best_b_indices     .empty else None
                    worst_idx   = elo_ser.loc[worst_b_indices]  .idxmax() if not worst_b_indices    .empty else None

                stats[col] = {"best_idx": best_idx, "worst_idx": worst_idx}

    borders = []

    if "GR" in df.columns:
        gv_series   = pd.to_numeric(df["GR"].astype(str).str.replace("%", ""), errors = "coerce")
        borders     = get_threshold_borders(analyzer, gv_series)

    col_borders = {"Player", "Score", "Mean Over-8", "Lives Saved", "IN Δ", "Rig Rate", "Solo Rig Rate", "Over-8 Δ", "Rig Δ", "Median Vintage Hit", "Metric", "Value", "Team Leader"}

    if "Score"  not in df.columns : col_borders.add("GR Δ") if "GR Δ" in df.columns else col_borders.add("GR")
    if "IN Δ"   not in df.columns : col_borders.add("IN GR")

    th_cells = []

    for c in df.columns:
        s_th = ' style="border-right: 3px solid black;"' if c in col_borders else ""
        th_cells.append(f"<th{s_th}>{str(c).replace(' ', '<br>')}</th>")

    html            = "<thead><tr>" + "".join(th_cells) + "</tr></thead><tbody>"
    bold_columns    = {"Player", "Metric", "Team Leader"}

    for idx, row in df.iterrows():
        b_s     =   "border-bottom: 3px solid black;" if idx in borders else ""
        html    +=  "<tr>"

        for cname, cell in row.items():
            style = [b_s] if b_s else []
            if cname in col_borders: style.append("border-right: 3px solid black;")

            if      cname == "Mean Difficulty Hit"  and pd.notnull(cell) and isinstance(cell, (int, float)) : cell_display = f"{float(cell):.2f}"
            elif    cname == "Median Vintage Hit"   and pd.notnull(cell) and isinstance(cell, (int, float)) : cell_display = format_year(float(cell))
            else                                                                                            : cell_display = cell

            if cname in stats:
                val_best_idx    = stats[cname]["best_idx"]
                val_worst_idx   = stats[cname]["worst_idx"]
                is_max          = idx == val_best_idx   if cname in desc else idx == val_worst_idx
                is_min          = idx == val_worst_idx  if cname in desc else idx == val_best_idx
                elig            = True if mask is None or cname not in rest else mask[idx]

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
                body                {{font-family: 'Segoe UI'; background: white; display: inline-block; margin: 0}}
                h2                  {{margin: 0 0 10px 0; font-size: 40px; text-align: center}}
                table               {{border-collapse: collapse; width: auto; border: 3px solid black}}
                th                  {{font-weight: bold; font-size: 25px; text-align: center; padding: 5px 10px; border: 1px solid black; border-bottom: 3px solid black; background-color: #f0f0f0}}
                td                  {{font-size: 25px; text-align: center; padding: 5px 10px; border: 1px solid black}}
                tr:nth-child(even)  {{background-color: #f0f0f0}}
            </style>
        </head>
        <body>
            <h2>{title}</h2>
            <table>{html}</table>
        </body>
    </html>"""

    hti = Html2Image(
        size                = (max(2000, len(df.columns) * 120), max(2000, len(df) * 60)),
        browser_executable  = analyzer.browser_path,
        output_path         = str(path),
        custom_flags        = ["--log-level=3", "--silent"],
    )

    hti.screenshot(html_str = full, save_as = fname)

    try                 : trim_whitespace(path / fname)
    except Exception    : pass

def fuse_images(path: Path):
    f   = {"Tour": "Tour.png", "Team": "Team.png", "Tier": "Tier.png", "Guess": "Guess.png", "List": "List.png", "Song": "Song.png"}
    ps  = {k: path / v for k, v in f.items() if (path / v).exists()}
    imgs= {k: Image.open(v) for k, v in ps.items()}

    if "Tour" not in imgs:
        for k, p in ps.items():
            if k not in ["List", "Guess", "Song"]:
                try                 : os.remove(p)
                except Exception    : pass

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

        try                 : trim_whitespace(plots_out_p)
        except Exception    : pass

        img_plots = Image.open(plots_out_p)

    left_components = [img_tour]

    if img_team: left_components.append(img_team)
    if img_tier: left_components.append(img_tier)

    gap_size    = 10
    total_gaps  = gap_size * (len(left_components) - 1)
    left_h_raw  = sum(img.height    for img in left_components) + total_gaps
    left_w_max  = max(img.width     for img in left_components)

    if img_plots:
        plots_aspect        = img_plots.width / img_plots.height
        plots_h_scaled      = left_h_raw
        plots_w_scaled      = int(plots_h_scaled * plots_aspect)
        img_plots_scaled    = img_plots.resize((plots_w_scaled, plots_h_scaled), Image.Resampling.LANCZOS)
    else: plots_w_scaled    = 0

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

    try                 : trim_whitespace(extra_out_p)
    except Exception    : pass

    img_extra   = Image.open(extra_out_p)
    p_path      = path / "Player.png"

    if p_path.exists():
        img_player          = Image.open(p_path)
        player_aspect       = img_player.width / img_player.height
        player_w_scaled     = img_extra.width
        player_h_scaled     = int(player_w_scaled / player_aspect)
        img_player_scaled   = img_player.resize((player_w_scaled, player_h_scaled), Image.Resampling.LANCZOS)
        gen_w               = img_extra.width
        gen_h               = player_h_scaled + 10 + img_extra.height
        img_general         = Image.new("RGB", (gen_w, gen_h), "white")

        img_general.paste(img_player_scaled,    (0, 0))
        img_general.paste(img_extra,            (0, player_h_scaled + 10))

        gen_out_p = path / "General.png"
        img_general.save(gen_out_p, compress_level = 9, optimize = True)

        try                 : trim_whitespace(gen_out_p)
        except Exception    : pass