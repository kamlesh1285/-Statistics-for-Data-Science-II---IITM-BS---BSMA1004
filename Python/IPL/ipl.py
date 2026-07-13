# Write a python script   1. To read all IPL yaml files and collect the actual data  you need.     2.  Write the data into spreadsheet   3. Several Python data Structures (list, dictionaries)  and modules (os, yaml, pandas)  will be useful.     Take IPL data from cricksheet.org


# How to use it:

# Get the data: my sandbox has no internet access, so download the IPL YAML zip yourself from https://cricsheet.org/downloads/ (look for the IPL YAML archive) and unzip it into a folder, e.g. ipl_yaml_data/.
# Install deps: pip install pyyaml pandas openpyxl
# Run:

 #  python ipl_yaml_to_excel.py --input ipl_yaml_data --output ipl_data.xlsx
# What it does:

# os/glob find every .yaml file in the input folder
# yaml.safe_load parses each file into nested dict/list structures
# For each match, it pulls match-level info (teams, venue, toss, winner, margin, player of the match, umpires) into a dict, and walks every innings/delivery into a list of ball-by-ball dicts
# pandas turns both lists of dicts into DataFrames and writes them to two sheets (Matches, Deliveries) in one .xlsx file via ExcelWriter

# It's built to tolerate minor format differences across seasons (e.g. old vs. new wicket format) and skips/logs any file it can't parse rather than crashing the whole run. Let me know if you'd like it extended — e.g. player career stats, team-vs-team win tallies, or a CSV output option instead of Excel.



"""
ipl_yaml_to_excel.py

Reads all Cricsheet IPL match YAML files from a local folder, extracts the
useful match-level and ball-by-ball data, and writes it into an Excel
workbook (two sheets: "Matches" and "Deliveries").

------------------------------------------------------------------------
STEP 1: GET THE DATA
------------------------------------------------------------------------
Cricsheet doesn't allow programmatic scraping in this environment, so
download the IPL YAML data manually first:

    1. Go to https://cricsheet.org/downloads/
    2. Find the "IPL" section and download the "yaml" zip
       (direct link is usually https://cricsheet.org/downloads/ipl_male_csv2.zip
        or similar - check the page for the current .zip link, since file
        names/versions change over time).
    3. Unzip it into a folder, e.g.:

         ipl_yaml_data/
             335982.yaml
             335983.yaml
             ...

------------------------------------------------------------------------
STEP 2: RUN THIS SCRIPT
------------------------------------------------------------------------
    pip install pyyaml pandas openpyxl
    python ipl_yaml_to_excel.py --input ipl_yaml_data --output ipl_data.xlsx

------------------------------------------------------------------------
WHAT DATA IS EXTRACTED
------------------------------------------------------------------------
Matches sheet (one row per match):
    match_id, date, season, venue, city, team1, team2,
    toss_winner, toss_decision, winner, win_by_type, win_by_margin,
    player_of_match, umpire1, umpire2

Deliveries sheet (one row per ball bowled):
    match_id, innings, batting_team, bowling_team, over, ball,
    batsman, bowler, non_striker, runs_batsman, runs_extras,
    runs_total, extra_type, wicket_kind, player_dismissed
"""

import os
import glob
import argparse
import yaml
import pandas as pd


def parse_match(filepath):
    """Parse a single Cricsheet YAML file and return (match_row, list_of_delivery_rows)."""
    with open(filepath, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    match_id = os.path.splitext(os.path.basename(filepath))[0]
    info = data.get("info", {})

    teams = info.get("teams", [None, None])
    team1 = teams[0] if len(teams) > 0 else None
    team2 = teams[1] if len(teams) > 1 else None

    outcome = info.get("outcome", {}) or {}
    winner = outcome.get("winner")
    by = outcome.get("by", {}) or {}
    win_by_type = "runs" if "runs" in by else ("wickets" if "wickets" in by else None)
    win_by_margin = by.get("runs", by.get("wickets"))

    toss = info.get("toss", {}) or {}
    dates = info.get("dates", [])
    umpires = info.get("umpires", [None, None])

    player_of_match = info.get("player_of_match", [None])
    player_of_match = player_of_match[0] if player_of_match else None

    match_row = {
        "match_id": match_id,
        "date": dates[0] if dates else None,
        "season": info.get("season"),
        "venue": info.get("venue"),
        "city": info.get("city"),
        "team1": team1,
        "team2": team2,
        "toss_winner": toss.get("winner"),
        "toss_decision": toss.get("decision"),
        "winner": winner,
        "win_by_type": win_by_type,
        "win_by_margin": win_by_margin,
        "player_of_match": player_of_match,
        "umpire1": umpires[0] if len(umpires) > 0 else None,
        "umpire2": umpires[1] if len(umpires) > 1 else None,
    }

    delivery_rows = []
    innings_list = data.get("innings", [])
    for inn_idx, inn in enumerate(innings_list, start=1):
        # each inn is a dict like {"1st innings": {"team": ..., "deliveries": [...]}}
        inn_name, inn_data = next(iter(inn.items()))
        batting_team = inn_data.get("team")
        bowling_team = team2 if batting_team == team1 else team1
        deliveries = inn_data.get("deliveries", [])

        for delivery in deliveries:
            # each delivery is {0.1: {batsman, bowler, non_striker, runs: {...}, ...}}
            over_ball, d = next(iter(delivery.items()))
            over = int(over_ball)
            ball = round((over_ball - over) * 100)  # e.g. 0.1 -> ball 1

            runs = d.get("runs", {}) or {}
            extras = d.get("extras", {}) or {}
            extra_type = next(iter(extras.keys()), None)

            wicket_kind = None
            player_dismissed = None
            wickets = d.get("wicket")
            if wickets:
                # older format: single dict; newer format: list of dicts
                wicket_entry = wickets[0] if isinstance(wickets, list) else wickets
                wicket_kind = wicket_entry.get("kind")
                player_dismissed = wicket_entry.get("player_out")

            delivery_rows.append({
                "match_id": match_id,
                "innings": inn_idx,
                "batting_team": batting_team,
                "bowling_team": bowling_team,
                "over": over,
                "ball": ball,
                "batsman": d.get("batsman"),
                "bowler": d.get("bowler"),
                "non_striker": d.get("non_striker"),
                "runs_batsman": runs.get("batsman", 0),
                "runs_extras": runs.get("extras", 0),
                "runs_total": runs.get("total", 0),
                "extra_type": extra_type,
                "wicket_kind": wicket_kind,
                "player_dismissed": player_dismissed,
            })

    return match_row, delivery_rows


def main():
    parser = argparse.ArgumentParser(description="Convert Cricsheet IPL YAML files to an Excel workbook.")
    parser.add_argument("--input", required=True, help="Folder containing IPL .yaml files")
    parser.add_argument("--output", default="ipl_data.xlsx", help="Output .xlsx file path")
    args = parser.parse_args()

    yaml_files = glob.glob(os.path.join(args.input, "*.yaml")) + glob.glob(os.path.join(args.input, "*.yml"))
    if not yaml_files:
        print(f"No YAML files found in '{args.input}'. Check the folder path.")
        return

    print(f"Found {len(yaml_files)} YAML files. Parsing...")

    all_matches = []
    all_deliveries = []

    for i, filepath in enumerate(sorted(yaml_files), start=1):
        try:
            match_row, delivery_rows = parse_match(filepath)
            all_matches.append(match_row)
            all_deliveries.extend(delivery_rows)
        except Exception as e:
            print(f"  Skipped {os.path.basename(filepath)}: {e}")

        if i % 50 == 0 or i == len(yaml_files):
            print(f"  Processed {i}/{len(yaml_files)} files")

    matches_df = pd.DataFrame(all_matches)
    deliveries_df = pd.DataFrame(all_deliveries)

    print(f"Writing {len(matches_df)} matches and {len(deliveries_df)} deliveries to '{args.output}'...")

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        matches_df.to_excel(writer, sheet_name="Matches", index=False)
        deliveries_df.to_excel(writer, sheet_name="Deliveries", index=False)

    print("Done.")


if __name__ == "__main__":
    main()