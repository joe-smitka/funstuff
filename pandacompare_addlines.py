import msoffcrypto
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

def decrypt_excel(file_path, password):
    decrypted = io.BytesIO()
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted

# === SETTINGS ===
password = "yourpassword"
file1_path = "file1.xlsx"
file2_path = "file2.xlsx"
output_path = "highlighted_changes.xlsx"
id_column = "ID"  # <-- update this to your actual ID column name

# === STEP 1: Decrypt both files ===
decrypted_file1 = decrypt_excel(file1_path, password)
decrypted_file2 = decrypt_excel(file2_path, password)

# === STEP 2: Load with pandas ===
df1 = pd.read_excel(decrypted_file1)
df2 = pd.read_excel(decrypted_file2)

# === STEP 3: Align on ID ===
df1.set_index(id_column, inplace=True)
df2.set_index(id_column, inplace=True)

# Reindex both to include all unique IDs from both files
all_ids = df1.index.union(df2.index)
df1 = df1.reindex(all_ids)
df2 = df2.reindex(all_ids)

# === STEP 4: Compare with null-safe fill ===
df1_filled = df1.fillna("__NA__")
df2_filled = df2.fillna("__NA__")
diff_df = df1_filled.ne(df2_filled)

# === STEP 5: Reload decrypted_file2 into openpyxl for highlighting ===
decrypted_file2.seek(0)
wb = load_workbook(filename=decrypted_file2)
ws = wb.active

# === STEP 6: Build row lookup: map ID to Excel row number ===
# Assumes header is on row 1
id_to_row = {}
for row in range(2, ws.max_row + 1):  # skip header row
    cell_value = ws.cell(row=row, column=1).value  # assumes ID is in col A (adjust if needed)
    if cell_value is not None:
        id_to_row[cell_value] = row

# === STEP 7: Highlighting logic ===
highlight_change = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
highlight_new = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

for idx, row in diff_df.iterrows():
    if idx not in id_to_row:
        continue  # New row not present in second file, skip (or handle differently)
    excel_row = id_to_row[idx]
    for col_idx, changed in enumerate(row):
        if changed:
            ws.cell(row=excel_row, column=col_idx + 2).fill = highlight_change  # +2: col A is ID

# === STEP 8: Add any new rows not found in df1 ===
new_rows = df2.index.difference(df1.index)
for new_id in new_rows:
    if new_id not in id_to_row:
        continue
    row_num = id_to_row[new_id]
    for col in range(1, len(df2.columns) + 1):
        ws.cell(row=row_num, column=col + 1).fill = highlight_new  # green fill for new row

# === STEP 9: Save result ===
wb.save(output_path)
print(f"Done! Highlighted changes saved to {output_path}")
