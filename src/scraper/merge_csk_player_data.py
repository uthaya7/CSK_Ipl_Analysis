import pandas as pd

# ---------- STEP 1: LOAD DATA ----------
# Change file paths to your actual Excel locations
batting = pd.read_excel("data/raw_temp/batting_records_csk.xlsx")
bowling = pd.read_excel("data/raw_temp/bowling_records_csk.xlsx")
fielding = pd.read_excel("data/raw_temp/fielding_records_csk.xlsx")

# ---------- STEP 2: CLEAN PLAYER NAMES ----------
batting["Player"] = batting["Player"].str.strip().str.title()
bowling["Player"] = bowling["Player"].str.strip().str.title()
fielding["Player"] = fielding["Player"].str.strip().str.title()

# ---------- STEP 3: MERGE DATA (SEQUENTIALLY) ----------
# 1. Merge Batting and Bowling DataFrames
# Outer join ensures all players are included
merged_df = pd.merge(
    batting,
    bowling,
    on="Player",
    how="outer", 
    suffixes=("_bat", "_bowl") # Differentiates columns from batting and bowling
)

# 2. Merge the result with the Fielding DataFrame
merged_df = pd.merge(
    merged_df,
    fielding,
    on="Player",
    how="outer" # Keep outer merge to ensure all players remain
    # No need for suffixes here unless 'fielding' has same column names as others, 
    # but Pandas will add '_x' and '_y' if needed.
)

# ---------- STEP 4: SORT & RESET INDEX ----------
merged_df = merged_df.sort_values(by="Player").reset_index(drop=True)

# ---------- STEP 5: SAVE TO EXCEL ----------
merged_df.to_excel("CSK_players_merged.xlsx", index=False)
print(f"✅ Merged file created successfully! Total players: {merged_df.shape[0]}")