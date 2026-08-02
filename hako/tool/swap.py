import glob, re, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def main():
    txt_files = glob.glob("*.txt")

    if not txt_files:
        print("No TXT file found")
        return

    file_path = txt_files[0]
    with open(file_path, 'r', encoding = 'utf-8') as f: lines = f.readlines()
        
    player_pattern  = re.compile(r'([a-zA-Z0-9_]+)\s*\(([\d.]+)\)')
    total_pattern   = re.compile(r'Total = ([\d.]+)')
    teams           = []
    
    for i, line in enumerate(lines):
        if 'Total =' in line:
            players     = player_pattern    .findall(line)
            total_match = total_pattern     .search(line)

            if players and total_match:
                teams.append({
                    'line_index'    : i,
                    'players'       : players,
                    'total'         : float(total_match.group(1))
                })

    if not teams:
        print("No team data found")
        return

    print("Which team do you want to swap a player for?")

    for i, team in enumerate(teams):
        first_player = team['players'][0][0]
        print(f"[{i + 1}] {first_player}")

    try:
        team_choice     = int(input("> ")) - 1
        selected_team   = teams[team_choice]

    except (ValueError, IndexError):
        print("Invalid team choice")
        return

    print("\nWhich player do you want to swap for?")

    for i, (player, _) in enumerate(selected_team['players']):
        print(f"[{i + 1}] {player}")

    try:
        player_choice           = int(input("> ")) - 1
        old_player, old_elo_str = selected_team['players'][player_choice]
        old_elo                 = float(old_elo_str)

    except (ValueError, IndexError):
        print("Invalid player choice")
        return

    print("\nWith whom?")
    new_player = input("Player: ").strip()

    print("\nWhat is their Elo?")

    try:
        new_elo_str = input("Elo: ").strip()
        new_elo     = float(new_elo_str)

    except ValueError:
        print("Invalid Elo")
        return
    
    line_idx    = selected_team['line_index']
    old_total   = selected_team['total']
    new_total   = old_total - old_elo + new_elo
    target_line = lines[line_idx]

    target_line = target_line.replace(
        f"{old_player} ({old_elo_str})", 
        f"{new_player.lower()} ({new_elo_str})"
    )

    target_line     = re.sub(r'Total = [\d.]+', f"Total = {new_total:.3f}", target_line)
    lines[line_idx] = target_line
    
    sum_totals = 0
    team_count = 0

    for line in lines:
        if 'Total =' in line:
            match = total_pattern.search(line)

            if match:
                sum_totals += float(match.group(1))
                team_count += 1

    new_average = sum_totals / team_count if team_count > 0 else 0

    for i, line in enumerate(lines):
        if 'Average:' in line:
            lines[i] = re.sub(r'Average:\s*[\d.]+', f"Average: {new_average:.3f}", line)
            break

    with open(file_path, 'w', encoding = 'utf-8') as f: f.writelines(lines)
    print(f"\n{old_player} has been swapped for {new_player.lower()}")

if __name__ == "__main__": main()