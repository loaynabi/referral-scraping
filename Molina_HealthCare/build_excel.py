"""
Builds Molina_HealthCare_Providers.xlsx from checkpoint.json.
Forces NPI / phone / zip columns to text format so leading digits/formatting
are not mangled by Excel's auto-number conversion.
"""
import json
import os
import pandas as pd
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE, "checkpoint.json"), encoding="utf-8") as f:
    records = json.load(f)

df = pd.DataFrame(records)

# Column order / friendly names
col_order = [
    "name", "primary_specialty", "npi", "medical_group",
    "address", "city", "state", "zip", "phone",
    "accepting_new_patients", "source_search_city", "source_search_zip",
]
df = df[[c for c in col_order if c in df.columns]]
df.columns = [
    "Provider Name", "Specialty", "NPI", "Medical Group / Location",
    "Address", "City", "State", "ZIP", "Phone",
    "Accepting New Patients (Y/N)", "Search Metro", "Search ZIP",
]

# Sort for readability
df = df.sort_values(["Search Metro", "City", "Provider Name"]).reset_index(drop=True)

out_path = os.path.join(BASE, "Molina_HealthCare_Providers.xlsx")
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Molina Medi-Cal PCPs")
    ws = writer.sheets["Molina Medi-Cal PCPs"]

    # Force text format on NPI, Phone, ZIP, Search ZIP columns
    text_cols = ["NPI", "Phone", "ZIP", "Search ZIP"]
    header = [cell.value for cell in ws[1]]
    for col_name in text_cols:
        if col_name not in header:
            continue
        col_idx = header.index(col_name) + 1
        col_letter = get_column_letter(col_idx)
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row}"]
            if cell.value is not None:
                cell.value = str(cell.value)
            cell.number_format = "@"

    # Autofit-ish column widths
    for i, col_name in enumerate(header, start=1):
        col_letter = get_column_letter(i)
        max_len = max(
            [len(str(col_name))] +
            [len(str(ws[f"{col_letter}{r}"].value or "")) for r in range(2, ws.max_row + 1)]
        )
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

    ws.freeze_panes = "A2"

print(f"Wrote {len(df)} rows to {out_path}")
