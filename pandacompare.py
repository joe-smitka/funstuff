import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import msoffcrypto
import io

# Decrypt file into memory
def decrypt_excel(file_path, password):
    decrypted = io.BytesIO()
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted

# Use pandas on the decrypted file
decrypted_file1 = decrypt_excel("file1.xlsx", "yourpassword")
decrypted_file2 = decrypt_excel("file2.xlsx", "yourpassword")

df1 = pd.read_excel(decrypted_file1)
df2 = pd.read_excel(decrypted_file2)

# Load both Excel files
#df1 = pd.read_excel("file1.xlsx")
#df2 = pd.read_excel("file2.xlsx")

# Create a difference dataframe
diff_df = df1.ne(df2)

# Load workbook with openpyxl for highlighting
wb = load_workbook("file2.xlsx")
ws = wb.active
highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Highlight changed cells
for row in range(2, len(df2) + 2):  # +2 because openpyxl rows are 1-indexed and header is row 1
    for col in range(1, len(df2.columns) + 1):
        if diff_df.iloc[row - 2, col - 1]:
            ws.cell(row=row, column=col).fill = highlight

# Save the new file
wb.save("highlighted_changes.xlsx")
