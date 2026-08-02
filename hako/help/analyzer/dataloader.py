import datetime
import os
import re
from collections import Counter, defaultdict
from pathlib import Path
from tkinter import messagebox

from dateutil.relativedelta import relativedelta
import pandas as pd

from hako.help.config import BROWSER_PATHS, FILE_ALIAS, FILE_CODES, TEAMS_RE, URL_ALIAS


def find_browser() -> str | None:
    """Locates the first available web browser executable from predefined system paths."""
    return next((p for p in BROWSER_PATHS if os.path.exists(p)), None)


def load_player_ids() -> dict[str, str]:
    """Loads player name-to-ID mapping from the online alias registry CSV."""
    id_map = {}
    try:
        df = pd.read_csv(URL_ALIAS)
        for _, row in df.iterrows():
            name = str(row.get("Player Name", "")).strip().lower()
            pid = str(row.get("Player ID", "")).strip()
            if name and pid:
                id_map[name] = pid
    except Exception:
        pass
    return id_map


def internal_clean_data(idtable: list, statstable: list, is_watched: bool) -> pd.DataFrame:
    """Parses and normalizes raw stats and ID tables retrieved from Google Sheets."""
    headers = idtable[0]
    data = idtable[1:]
    alias_df = pd.DataFrame(data, columns=headers)
    alias_df["Player Name"] = alias_df["Player Name"].str.strip().str.lower()
    alias_to_id = dict(zip(alias_df["Player Name"], alias_df["Player ID"]))

    headers = statstable[0]
    data = statstable[1:]
    df = pd.DataFrame(data, columns=headers)
    df = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")

    df["Player ID"] = df["Player name"].dropna().str.strip().str.lower().map(alias_to_id)
    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")

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
        "IN guess rate",
    ]

    watched_cols = [
        "Rigs hit",
        "Rigs",
        "Rigs missed",
        "Solo rigs",
        "Missed solos",
        "Lives lost on rigs",
        "Offlist erigs",
        "avg/8 of your rigs",
    ]

    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")

    if is_watched:
        df[watched_cols] = df[watched_cols].apply(pd.to_numeric, errors="coerce")
        cols.extend(watched_cols)
        df["Offlist hit"] = df["Total hit"] - df["Rigs hit"]

    df = df[
        (
            pd.to_numeric(df["WIN"], errors="coerce").fillna(0)
            + pd.to_numeric(df["LOSE"], errors="coerce").fillna(0)
            + pd.to_numeric(df["TIE"], errors="coerce").fillna(0)
        )
        >= 4
    ]

    return df


def clean_data_local(
    idtable: list, statstable: list, max_fallback_window: int, active_tours: int, is_list: bool
) -> pd.DataFrame:
    """Filters, aggregates, and computes historic baselines for player performance."""
    df = internal_clean_data(idtable, statstable, is_list)
    fallback_cols = ["Player ID", "Guess rate", "Usefulness", "OP guess rate", "ED guess rate", "IN guess rate"]

    if df.empty:
        return pd.DataFrame(columns=fallback_cols)

    six_months_ago = datetime.datetime.now() - relativedelta(months=max_fallback_window)
    year_6m_ago = six_months_ago.year
    month_6m_ago = six_months_ago.month

    year_df = df[
        ((df["Timestamp"].dt.year > year_6m_ago))
        | ((df["Timestamp"].dt.year == year_6m_ago) & (df["Timestamp"].dt.month >= month_6m_ago))
    ]
    if year_df.empty:
        return pd.DataFrame(columns=fallback_cols)

    year_df = year_df.sort_values(["Player ID", "Timestamp"])
    result_df = year_df.groupby("Player ID").tail(active_tours)

    gr_cols = ["Guess rate", "Usefulness", "OP guess rate", "ED guess rate", "IN guess rate"]
    agg_dict = {col: "mean" if col in gr_cols else "max" for col in result_df.columns if col != "Player ID"}
    agg_dict = {k: v for k, v in agg_dict.items() if k in result_df.columns}

    result_df = result_df.groupby("Player ID").agg(agg_dict).reset_index()
    result_df["Player ID"] = result_df["Player ID"].astype(int)

    return result_df


def generate_acronyms(active_names: set[str] | list[str]) -> dict[str, str]:
    """Generates unique 3+ character acronyms for a set of player display names."""
    acronyms = {}
    for name in active_names:
        clean = "".join(filter(str.isalnum, name))
        length = 3
        acr = clean[:length].upper() if len(clean) >= length else clean.upper().ljust(length, "X")
        acronyms[name.lower()] = acr

    while True:
        counts = Counter(acronyms.values())
        duplicates = {acr for acr, count in counts.items() if count > 1}

        if not duplicates:
            break

        for name in active_names:
            n_low = name.lower()
            if acronyms[n_low] in duplicates:
                clean = "".join(filter(str.isalnum, name))
                curr_len = len(acronyms[n_low])
                next_len = curr_len + 1

                if next_len <= len(clean):
                    acronyms[n_low] = clean[:next_len].upper()
                else:
                    acronyms[n_low] = clean.upper() + str(next_len - len(clean))

    return acronyms


def scan_players(paths: list[Path]) -> tuple[set[str], defaultdict[str, set[str]]]:
    """Scans all tournament match JSON files to identify player display names and appearances."""
    import json

    players = set()
    apps = defaultdict(set)

    for p in paths:
        try:
            with open(p, encoding="utf-8") as f:
                data = json.load(f)
                for s in data.get("songs", []):
                    for plyr in s.get("correctGuessPlayers", []):
                        if isinstance(plyr, str):
                            players.add(plyr)
                            apps[plyr].add(str(p))
                        elif isinstance(plyr, dict) and "name" in plyr:
                            players.add(plyr["name"])
                            apps[plyr["name"]].add(str(p))

                    for ls in s.get("listStates", []):
                        players.add(ls["name"])
                        apps[ls["name"]].add(str(p))
        except Exception:
            continue

    return players, apps


def load_team_data(
    tour_dir: Path,
    all_known: set[str],
    id_database: dict[str, str],
    subbed_players_set: set[str],
    main_roster_names: set[str],
) -> tuple[bool, dict, dict, dict, defaultdict, set, list, list]:
    """Parses code.txt team definitions and aligns player names against alias configurations."""
    codes = tour_dir / FILE_CODES
    if not codes.exists() or os.path.getsize(codes) == 0:
        return False, {}, {}, {}, defaultdict(set), all_known, [], list(all_known)

    with open(codes, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    has_avg = any(l.lower().startswith(("average", "avg")) for l in lines)
    team_lines = 0
    bad_lines = []

    for line in lines:
        if line.lower().startswith(("average", "avg", "sub")) or line.startswith("http"):
            continue
        team_text = line.split("|", 1)[0].strip()

        if not re.findall(TEAMS_RE, team_text):
            bad_lines.append(line)
        else:
            team_lines += 1

    if bad_lines or team_lines == 0 or not has_avg:
        error_details = ["Broken codes.txt"]
        if bad_lines:
            error_details.append(f"Cannot parse line structure: '{bad_lines[0]}'")
        if not has_avg:
            error_details.append("Missing (Average: Elo) line")

        messagebox.showerror("Invalid Code File", "\n".join(error_details))
        return False, {}, {}, {}, defaultdict(set), all_known, [], list(all_known)

    elo_map = {}
    assignments = {}
    rosters = defaultdict(set)
    t1_lookup = {}
    avail = sorted(list(all_known))
    alias_path = tour_dir / FILE_ALIAS
    local_aliases = {}

    if alias_path.exists():
        with open(alias_path, "r", encoding="utf-8") as f:
            for line in f:
                if "," in line:
                    k, v = line.strip().split(",", 1)
                    local_aliases[k.strip().lower()] = v.strip()

    new_aliases = {}

    def find_best_match(p_in: str) -> str | None:
        p_low = p_in.lower()
        if p_low in local_aliases:
            m = local_aliases[p_low]
            if m in all_known:
                return m

        match = next((n for n in all_known if n.lower() == p_low), None)
        if not match:
            if not id_database:
                id_database.update(load_player_ids())

            if p_low in id_database:
                target_id = id_database[p_low]
                match = next((n for n in all_known if id_database.get(n.lower()) == target_id), None)

        if match:
            new_aliases[p_in] = match
        return match

    for line in lines:
        matches = re.findall(TEAMS_RE, line)
        for p_in, val in matches:
            if not line.lower().startswith("subs:"):
                match = find_best_match(p_in)
                if match:
                    elo_map[match.lower()] = val

    idx = 1
    sub_candidates_raw = []

    for line in lines:
        if line.lower().startswith("subs:"):
            mems_subs = re.findall(TEAMS_RE, line)
            for p_sub, val_s in mems_subs:
                m_sub = find_best_match(p_sub)
                if m_sub:
                    subbed_players_set.add(m_sub.lower())
                    sub_candidates_raw.append(m_sub)
                    elo_map[m_sub.lower()] = val_s
            continue

        mems = re.findall(TEAMS_RE, line.split("|")[0])
        if not mems:
            continue

        p_captain, _ = mems[0]
        c_match = find_best_match(p_captain)
        ename = c_match if c_match else p_captain
        t1_lookup[idx] = ename

        for i, (p_in, _) in enumerate(mems[:4]):
            tier = str(i + 1)
            match = find_best_match(p_in)
            if match:
                main_roster_names.add(match.lower())
                assignments[match.lower()] = (idx, tier)
                rosters[idx].add(match)
                if match in avail:
                    avail.remove(match)

        idx += 1

    original_players_display = [
        p
        for p in all_known
        if p.lower() in assignments and p.lower() not in [s.lower() for s in sub_candidates_raw]
    ]

    if new_aliases:
        existing_pairs = set()
        if alias_path.exists():
            with open(alias_path, "r", encoding="utf-8") as f:
                for l in f:
                    parts = [p.strip().lower() for p in l.split(",")]
                    if len(parts) >= 2:
                        existing_pairs.add((parts[0], parts[1]))

        with open(alias_path, "a", encoding="utf-8") as f:
            for k, v in new_aliases.items():
                k_low, v_low = k.strip().lower(), v.strip().lower()
                if (k_low, v_low) not in existing_pairs:
                    f.write(f"{k}, {v}\n")
                    existing_pairs.add((k_low, v_low))

    return (
        True,
        elo_map,
        assignments,
        t1_lookup,
        rosters,
        all_known,
        sub_candidates_raw,
        original_players_display,
    )