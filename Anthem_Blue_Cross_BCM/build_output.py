"""
Parses raw scraped provider-card text blocks (collected from the Anthem Find Care
results UI via browser automation) into structured rows, writes a resumable
checkpoint.json, deduplicates, and writes the final Excel deliverable.

Each raw batch file in raw/*.json is a JSON array-of-arrays: one sub-array per
result page, each containing ~20 raw innerText strings from
`app-fc-provider-card-cmp` elements. Each card's innerText follows the pattern:

    <Network Status>
    <Name>
    <Specialty>
    <Street Address>
    <City, ST ZIP (distance mi)>
    <Phone>

(A few "Other Services" entries only have 5 lines, no distance.)
"""
import json
import re
import os
import glob

BASE_DIR = r"C:\Users\DamonHollister\Desktop\referral scraping\Anthem_Blue_Cross_BCM"
RAW_DIR = os.path.join(BASE_DIR, "raw")
CHECKPOINT_PATH = os.path.join(BASE_DIR, "checkpoint.json")
XLSX_PATH = os.path.join(BASE_DIR, "Anthem_Blue_Cross_Providers.xlsx")

CITY_STATE_ZIP_RE = re.compile(
    r"^(?P<city>.+),\s*(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})(?:-\d{4})?\s*(?:\((?P<distance>[\d.]+)\s*mi\))?\s*$"
)

# SSN / DOB guard patterns -- if any of these ever appear in scraped text, the
# field is dropped rather than included in output (per project safety rules).
SENSITIVE_PATTERNS = [
    re.compile(r"\bSSN\b", re.I),
    re.compile(r"\bsocial security\b", re.I),
    re.compile(r"\bdate of birth\b", re.I),
    re.compile(r"\bDOB\b"),
    re.compile(r"\b\d{3}-\d{2}-\d{4}\b"),  # SSN-shaped
]


def contains_sensitive(text: str) -> bool:
    return any(p.search(text) for p in SENSITIVE_PATTERNS)


def parse_card(card_text: str, source_zip: str):
    lines = [l.strip() for l in card_text.split("\n") if l.strip()]
    if not lines:
        return None
    if contains_sensitive(card_text):
        # Drop the field defensively; in practice this directory never
        # exposes SSN/DOB, but we guard per the project's safety rules.
        return None

    network_status = None
    idx = 0
    if lines[0] in ("In-Network", "Out-of-Network", "In-Network*"):
        network_status = lines[0]
        idx = 1

    name = lines[idx] if idx < len(lines) else ""
    specialty = lines[idx + 1] if idx + 1 < len(lines) else ""
    street = lines[idx + 2] if idx + 2 < len(lines) else ""
    city_state_zip_line = lines[idx + 3] if idx + 3 < len(lines) else ""
    phone = lines[idx + 4] if idx + 4 < len(lines) else ""

    city = state = zip_code = distance = ""
    m = CITY_STATE_ZIP_RE.match(city_state_zip_line)
    if m:
        city = m.group("city")
        state = m.group("state")
        zip_code = m.group("zip")
        distance = m.group("distance") or ""

    # Normalize phone to digits-preserving string (kept as text, not number)
    phone = phone.strip()

    row = {
        "Provider Type": "Primary Care",  # per task: category filter used for every search
        "Name": name,
        "Specialty": specialty,
        "Network Status": network_status or "",
        "Street Address": street,
        "City": city,
        "State": state,
        "Zip": zip_code,
        "Distance (mi)": distance,
        "Phone": phone,
        "Search Zip": source_zip,
        "Source": "Anthem Blue Cross CA Find Care - Medi-Cal HMO",
    }
    return row


def zip_from_filename(fname: str) -> str:
    base = os.path.basename(fname)
    m = re.match(r"(\d{5})_", base)
    return m.group(1) if m else "unknown"


def load_all_rows():
    rows = []
    for fpath in sorted(glob.glob(os.path.join(RAW_DIR, "*.json"))):
        zip_code = zip_from_filename(fpath)
        with open(fpath, "r", encoding="utf-8") as f:
            pages = json.load(f)
        for page in pages:
            for card_text in page:
                row = parse_card(card_text, zip_code)
                if row:
                    rows.append(row)
    return rows


def dedupe(rows):
    seen = {}
    for r in rows:
        key = (
            r["Name"].strip().lower(),
            r["Street Address"].strip().lower(),
            r["Specialty"].strip().lower(),
        )
        if key not in seen:
            seen[key] = r
    return list(seen.values())


def main():
    rows = load_all_rows()
    print(f"Parsed {len(rows)} raw rows from {RAW_DIR}")

    deduped = dedupe(rows)
    print(f"{len(deduped)} unique rows after dedup on (Name+Address+Specialty)")

    # --- checkpoint.json ---
    checkpoint = {
        "rows": deduped,
        "raw_row_count": len(rows),
        "unique_row_count": len(deduped),
        "zips_covered": sorted(set(r["Search Zip"] for r in deduped)),
        "resume_note": (
            "90012 (Los Angeles, downtown) covered through page 8 of the "
            "Medi-Cal HMO / Primary Care search (~160 raw cards, ~20/page). "
            "No other ZIP codes were completed in this run due to Anthem "
            "Find Care UI/session instability encountered mid-session -- see "
            "scrape_anthem.py notes. To resume: rerun the flow for 90012 "
            "starting at page 9, then repeat for 94102, 93701, 92101, etc."
        ),
    }
    with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
        json.dump(checkpoint, f, indent=2, ensure_ascii=False)
    print(f"Wrote checkpoint to {CHECKPOINT_PATH}")

    # --- Excel ---
    import pandas as pd

    columns = [
        "Provider Type", "Name", "Specialty", "Network Status",
        "Street Address", "City", "State", "Zip", "Distance (mi)",
        "Phone", "Search Zip", "Source",
    ]
    df = pd.DataFrame(deduped, columns=columns)
    df.sort_values(by=["Search Zip", "City", "Name"], inplace=True)

    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Providers")
        ws = writer.sheets["Providers"]
        # Force Phone and Zip columns to text format so Excel doesn't mangle
        # leading zeros / dashes / parens.
        phone_col = columns.index("Phone") + 1
        zip_col = columns.index("Zip") + 1
        search_zip_col = columns.index("Search Zip") + 1
        for col_idx in (phone_col, zip_col, search_zip_col):
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.number_format = "@"
                if cell.value is not None:
                    cell.value = str(cell.value)
        # Reasonable column widths
        widths = {
            "A": 14, "B": 28, "C": 34, "D": 14, "E": 30, "F": 16,
            "G": 7, "H": 10, "I": 12, "J": 16, "K": 12, "L": 38,
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    print(f"Wrote Excel file to {XLSX_PATH}")


if __name__ == "__main__":
    main()
