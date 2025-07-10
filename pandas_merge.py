import pandas as pd
from datetime import datetime, timedelta
import os
import openpyxl

# === CONFIGURATION ===
FOLDER = '.'  # Path to directory where the files live
INDEX_COL = 'Sys ID'
START_COL = 'BO'  # Excel column ID where enrichment begins

# === DATE LOGIC ===
today = datetime.today()
today_str = today.strftime('%y%m%d')
yesterday_str = (today - timedelta(days=1)).strftime('%y%m%d')

today_filename = f"{today_str}_Manifest_BCD.xlsx"
yesterday_filename = f"{yesterday_str}_Manifest_BCD.xlsx"
output_filename = f"{today_str}_Manifest_CTD.xlsx"

today_path = os.path.join(FOLDER, today_filename)
yesterday_path = os.path.join(FOLDER, yesterday_filename)
output_path = os.path.join(FOLDER, output_filename)

# === LOAD FILES ===
df_today = pd.read_excel(today_path, dtype=str)
df_yesterday = pd.read_excel(yesterday_path, dtype=str)

# Drop rows with no Sys ID
df_today = df_today[df_today[INDEX_COL].notna()]
df_yesterday = df_yesterday[df_yesterday[INDEX_COL].notna()]

# === IDENTIFY ENRICHED COLUMNS BY POSITION ===
def col_letter_to_index(letter):
    index = 0
    for char in letter:
        index = index * 26 + (ord(char.upper()) - ord('A')) + 1
    return index - 1  # Convert to 0-based

start_col_idx = col_letter_to_index(START_COL)

# Get list of columns from BO onward
enriched_columns = df_yesterday.columns[start_col_idx:]

# Pull only the enriched columns and Sys ID
df_extra = df_yesterday[[INDEX_COL] + list(enriched_columns)]

# === MERGE ===
df_merged = pd.merge(df_today, df_extra, on=INDEX_COL, how='left', suffixes=('', '_y'))

# Overwrite BO+ columns in today’s file with values from yesterday
for col in enriched_columns:
    merged_col = col + '_y'
    if merged_col in df_merged.columns:
        df_merged[col] = df_merged[merged_col]
        df_merged.drop(columns=[merged_col], inplace=True)

# === SAVE TO CTD FILE ===
df_merged.to_excel(output_path, index=False)
print(f"✅ Enrichment complete! Output saved to: {output_filename}")
