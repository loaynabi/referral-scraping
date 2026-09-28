import json
import pandas as pd

CHECKPOINT_PATH = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\checkpoint.json"
OUTPUT_XLSX = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\Health_Plan_of_San_Joaquin_Providers.xlsx"

with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
    data = json.load(f)
rows = data['rows']
print("raw rows:", len(rows))

df = pd.DataFrame(rows)

# Dedup on (NPI, Address, Specialty)
before = len(df)
df = df.drop_duplicates(subset=['NPI', 'Address', 'Specialty'], keep='first')
after = len(df)
print(f"deduped {before} -> {after} rows")

# Column order
cols = [
    'Provider Type', 'Name', 'Gender', 'Specialty', 'Office/Facility Name',
    'Address', 'Street', 'City', 'State', 'Zip',
    'Phone', 'Board Certified', 'Office Hours', 'Languages', 'Hospital',
    'NPI', 'License', 'Provider ID', 'Access', 'Accepting New Patients',
    'Age Group', 'Telehealth', 'Source Page (0-indexed)', 'Source Column',
]
cols = [c for c in cols if c in df.columns]
df = df[cols]

df = df.sort_values(['City', 'Specialty', 'Name']).reset_index(drop=True)

text_cols = ['Phone', 'NPI', 'Zip', 'License', 'Provider ID']

with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Primary Care Providers')
    ws = writer.sheets['Primary Care Providers']
    header = [cell.value for cell in ws[1]]
    for col_name in text_cols:
        if col_name not in header:
            continue
        col_idx = header.index(col_name) + 1
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(col_idx)
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row}"]
            cell.number_format = '@'
            if cell.value is not None:
                cell.value = str(cell.value)
    # auto width (rough)
    for i, col_name in enumerate(header, start=1):
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(i)
        maxlen = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).tolist()[:2000]])
        ws.column_dimensions[col_letter].width = min(max(maxlen + 2, 10), 45)

print("Wrote", OUTPUT_XLSX, "rows:", len(df))
