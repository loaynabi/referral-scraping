"""
HPSJ/MVHP Medicare Advantage D-SNP (H5734) Primary Care provider directory extractor,
plus merge with the existing Medi-Cal Primary Care Excel export into one unified file.

Part 1: Extract PCP rows from 3 county D-SNP PDFs (El Dorado/Alpine, San Joaquin,
        Stanislaus), using extract_words() + x0-column banding + label-anchored parsing.
        Resumable via checkpoint_dsnp.json.
Part 2: Load Health_Plan_of_San_Joaquin_Providers.xlsx (Medi-Cal Combined Directory),
        tag it with Directory Source, combine with the D-SNP rows, dedupe
        (NPI+Address+Specialty, falling back to Name+Address+Specialty), and write
        the merged workbook.
"""
import pdfplumber
import re
import json
import os
import time
import pandas as pd
from openpyxl.utils import get_column_letter

BASE = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin"
PDF_DIR = os.path.join(BASE, "pdf_county")
CHECKPOINT_PATH = os.path.join(BASE, "checkpoint_dsnp.json")
MEDICAL_XLSX = os.path.join(BASE, "Health_Plan_of_San_Joaquin_Providers.xlsx")
OUTPUT_XLSX = os.path.join(BASE, "HPSJ_MVHP_Merged_AllSources_Providers.xlsx")

DSNP_SOURCES = [
    {
        "key": "El_Dorado_Alpine",
        "path": os.path.join(PDF_DIR, "El_Dorado_and_Alpine_Directory_September_2026.pdf"),
        "county": "El Dorado/Alpine",
        "directory_source": "Medicare D-SNP - El Dorado/Alpine County",
        "pcp_start": 74,   # 0-indexed pdfplumber, inclusive (verified: header "C1. Primary Care Providers")
        "pcp_end": 79,     # 0-indexed pdfplumber, exclusive (verified: header "C2. Specialists" begins here)
    },
    {
        "key": "San_Joaquin",
        "path": os.path.join(PDF_DIR, "San_Joaquin_Directory_September_2026.pdf"),
        "county": "San Joaquin",
        "directory_source": "Medicare D-SNP - San Joaquin County",
        "pcp_start": 74,
        "pcp_end": 99,
    },
    {
        "key": "Stanislaus",
        "path": os.path.join(PDF_DIR, "Stanislaus_Directory_September_2026.pdf"),
        "county": "Stanislaus",
        "directory_source": "Medicare D-SNP - Stanislaus County",
        "pcp_start": 74,
        "pcp_end": 86,
    },
]

COL_BOUNDS = [(0, 175), (175, 385), (385, 999)]
HEADER_TOP_CUTOFF = 50  # lines above this 'top' are the repeated section-title banner; skip

GENDER_LICENSE_RE = re.compile(r'^(Male|Female)\s*\|\s*License:\s*(.+)$', re.I)
SPECIALTY_RE = re.compile(r'^Specialty:\s*(.+)$', re.I)
NPI_RE = re.compile(r'^NPI:\s*(.+)$', re.I)
BOARD_CERT_RE = re.compile(r'^Board Certified:\s*([YN])\s*$', re.I)
CITYSTATEZIP_RE = re.compile(r'^(?P<city>.+?),\s*(?P<state>[A-Z]{2})\s+(?P<zip>[\d\-]+)\s*$')
PHONE_RE = re.compile(r'^\d{3}-\d{3}-\d{4}')
HOURS_RE = re.compile(r'(\d{1,2}(:\d{2})?\s*(am|pm)|^(M|T|W|Th|F|Sa|Su)(-[A-Za-z]{1,2})?:|Telehealth\s*Only|24\s*hours|By\s*Appointment)', re.I)
ACCESS_CODE_RE = re.compile(r'^([A-Z]{1,3}(,\s*[A-Z]{1,3})*)$')
HAS_ALNUM_RE = re.compile(r'[A-Za-z0-9]')
ADDR_LIKE_RE = re.compile(r'^\d+\s')
SENSITIVE_RE = re.compile(r'\bSSN\b|\bsocial security\b|\bDOB\b|\bdate of birth\b', re.I)


def clean_value(v):
    if not isinstance(v, str):
        return v
    v = re.sub(r'\s+', ' ', v).strip()
    return v


def get_column_lines(page, x_min, x_max):
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    col_words = [w for w in words if x_min <= w['x0'] < x_max]
    rows = {}
    for w in col_words:
        key = round(w['top'])
        rows.setdefault(key, []).append(w)
    lines = []
    for top in sorted(rows.keys()):
        if top < HEADER_TOP_CUTOFF:
            continue
        ws = sorted(rows[top], key=lambda w: w['x0'])
        text = " ".join(w['text'] for w in ws)
        if not HAS_ALNUM_RE.search(text):
            continue  # decorative icon-only glyph line
        lines.append(text.strip())
    return lines


def split_name_credential(name_line):
    parts = [p.strip() for p in name_line.split(',')]
    if len(parts) >= 2:
        credential = parts[-1]
        name = ", ".join(parts[:-1])
    else:
        credential = ""
        name = name_line
    return name, credential


def parse_column(lines, page_num, column, county, directory_source):
    rows = []
    n = len(lines)
    spec_idxs = [i for i, t in enumerate(lines) if SPECIALTY_RE.match(t)]

    for spec_idx in spec_idxs:
        specialty = SPECIALTY_RE.match(lines[spec_idx]).group(1).strip()

        # Walk backward to find the name line, skipping an optional Gender|License line.
        gender = ""
        license_no = ""
        j = spec_idx - 1
        if j >= 0 and GENDER_LICENSE_RE.match(lines[j]):
            m = GENDER_LICENSE_RE.match(lines[j])
            gender = m.group(1).strip()
            license_no = m.group(2).strip()
            j -= 1
        if j < 0:
            continue
        name_line = lines[j]
        # Defensive: name line shouldn't itself be a label line
        if ':' in name_line:
            continue
        name, credential = split_name_credential(name_line)

        # org / facility name is normally the line right after Specialty:, but some
        # entries omit the org line and go straight to the street address (detected
        # by a leading street number combined with a city/state/zip line following).
        org_idx = spec_idx + 1
        if org_idx >= n:
            continue
        candidate = lines[org_idx].strip()
        next_line = lines[org_idx + 1].strip() if org_idx + 1 < n else ""
        if ADDR_LIKE_RE.match(candidate) and (CITYSTATEZIP_RE.match(next_line) or 'telehealth' in candidate.lower()):
            org = ""
            addr_start_idx = org_idx
        else:
            org = candidate
            addr_start_idx = org_idx + 1
            # NOTE: org names occasionally wrap onto a second line that also carries
            # part of the street address on the same physical line (PDF word-wrap
            # artifact, e.g. "Family First Medical Care, A" / "Professional
            # Corporation, 1317 Oakdale Rd Ste 230"). We deliberately do NOT try to
            # merge a continuation line into org here: the ambiguity between "org
            # wrapped a line" vs "this is the street line" isn't reliably resolvable
            # from text alone, and a wrong guess risks losing the street address
            # entirely. Left as a known minor limitation (org name may be truncated
            # in these rare wrap cases; full address text, including any org-name
            # remainder, is still preserved in the Address/Street field).

        # Find NPI line (search forward from addr_start_idx)
        npi_idx = None
        for k in range(addr_start_idx, n):
            if NPI_RE.match(lines[k]):
                npi_idx = k
                break
        if npi_idx is None:
            continue  # malformed / truncated entry, skip

        address_lines = [lines[k].strip() for k in range(addr_start_idx, npi_idx) if lines[k].strip()]
        address_full = ", ".join(address_lines)
        street, city, state, zipc = "", "", "", ""
        if address_lines:
            last = address_lines[-1]
            czm = CITYSTATEZIP_RE.match(last)
            if czm:
                city = czm.group('city').strip()
                state = czm.group('state').strip()
                zipc = czm.group('zip').strip()
                street = ", ".join(address_lines[:-1]).strip()
            else:
                street = address_full

        npi = NPI_RE.match(lines[npi_idx]).group(1).strip()

        # Find Board Certified line (end anchor), search forward from npi_idx+1
        bc_idx = None
        for k in range(npi_idx + 1, n):
            if BOARD_CERT_RE.match(lines[k]):
                bc_idx = k
                break
        if bc_idx is None:
            continue
        board_certified = BOARD_CERT_RE.match(lines[bc_idx]).group(1).upper()

        middle_lines = [lines[k].strip() for k in range(npi_idx + 1, bc_idx) if lines[k].strip()]
        phone = ""
        hours_parts = []
        lang_parts = []
        access_parts = []
        for idx_m, ml in enumerate(middle_lines):
            if idx_m == 0 and PHONE_RE.match(ml):
                phone = ml
                continue
            if ACCESS_CODE_RE.match(ml.replace(' ', '')) and ',' in ml or (ACCESS_CODE_RE.match(ml) and len(ml) <= 25):
                # short comma-separated uppercase letter codes -> access codes
                access_parts.append(ml)
            elif HOURS_RE.search(ml):
                hours_parts.append(ml)
            else:
                lang_parts.append(ml)

        def _tidy_join(parts, sep):
            cleaned = [p.strip().rstrip(';,').strip() for p in parts]
            cleaned = [p for p in cleaned if p]
            return sep.join(cleaned)

        office_hours = _tidy_join(hours_parts, "; ")
        languages = _tidy_join(lang_parts, ", ")
        access_codes = _tidy_join(access_parts, ", ")

        row = {
            'Provider Type': 'Primary Care',
            'Name': name,
            'Credential': credential,
            'Gender': gender,
            'Specialty': specialty,
            'Office/Facility Name': org,
            'Address': address_full,
            'Street': street,
            'City': city,
            'State': state,
            'Zip': zipc,
            'Phone': phone,
            'Board Certified': board_certified,
            'Office Hours': office_hours,
            'Languages': languages,
            'NPI': npi,
            'License': license_no,
            'Access': access_codes,
            'County': county,
            'Directory Source': directory_source,
            'Source Page (0-indexed)': page_num,
            'Source Column': column,
        }
        row = {k: clean_value(v) for k, v in row.items()}

        blob = json.dumps(row)
        if SENSITIVE_RE.search(blob):
            continue

        rows.append(row)
    return rows


def parse_page(page, page_num, county, directory_source):
    rows = []
    for ci, (xmin, xmax) in enumerate(COL_BOUNDS):
        lines = get_column_lines(page, xmin, xmax)
        rows += parse_column(lines, page_num, ci, county, directory_source)
    return rows


def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"rows": [], "completed_sources": []}


def save_checkpoint(data):
    tmp = CHECKPOINT_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f)
    os.replace(tmp, CHECKPOINT_PATH)


def extract_all_dsnp():
    data = load_checkpoint()
    rows = data.get("rows", [])
    completed = set(data.get("completed_sources", []))

    for src in DSNP_SOURCES:
        if src["key"] in completed:
            print(f"Skipping {src['key']} (already completed per checkpoint)")
            continue
        print(f"=== Processing {src['key']} ({src['path']}) ===")
        # drop any partial rows from a previous incomplete run of this source
        rows = [r for r in rows if r.get('_source_key') != src['key']]
        t0 = time.time()
        with pdfplumber.open(src["path"]) as pdf:
            for i in range(src["pcp_start"], src["pcp_end"]):
                page = pdf.pages[i]
                try:
                    page_rows = parse_page(page, i, src["county"], src["directory_source"])
                except Exception as e:
                    print(f"  ERROR on page {i}: {e}")
                    page_rows = []
                for r in page_rows:
                    r['_source_key'] = src['key']
                rows.extend(page_rows)
                print(f"  page {i}: {len(page_rows)} rows")
        completed.add(src["key"])
        save_checkpoint({"rows": rows, "completed_sources": sorted(completed)})
        elapsed = time.time() - t0
        n_this = len([r for r in rows if r.get('_source_key') == src['key']])
        print(f"Checkpointed after {src['key']}: {n_this} rows this source, {len(rows)} total, {elapsed:.1f}s")

    for r in rows:
        r.pop('_source_key', None)
    return rows


def normalize_for_match(text):
    """Loose normalization used ONLY for building the dedup key (never written to
    output). Handles two real cross-source formatting differences observed between
    the pre-existing Medi-Cal export and the freshly-parsed D-SNP rows:
      - The Medi-Cal export stores zip as a bogus "XXXXX-XXXX" value where the
        4-digit suffix duplicates the last 4 digits of the 5-digit zip (a parsing
        artifact from the original scrape, confirmed on 91/91 distinct zip values
        in that file) while the D-SNP rows use a plain 5-digit zip.
      - Case differs (e.g. "FAMILY MEDICINE" vs "Family Medicine", "STOCKTON" vs
        "Stockton") and punctuation/whitespace varies slightly.
    """
    t = (text or '').strip().lower()
    t = re.sub(r'(\d{5})-\d{4}\b', r'\1', t)  # drop bogus/zip+4 suffix
    t = re.sub(r'[.,]', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def dedup_key_npi(row):
    npi = (row.get('NPI') or '').strip()
    if not npi:
        return None
    return (npi, normalize_for_match(row.get('Address')), normalize_for_match(row.get('Specialty')))


def dedup_key_fallback(row):
    return (
        normalize_for_match(row.get('Name')),
        normalize_for_match(row.get('Address')),
        normalize_for_match(row.get('Specialty')),
    )


def merge_and_dedupe(dsnp_rows):
    med_df = pd.read_excel(MEDICAL_XLSX)
    med_rows = med_df.to_dict('records')
    for r in med_rows:
        r['Directory Source'] = 'Medi-Cal Combined Directory (HPSJ-MVHP)'
        for k in ('NPI', 'License', 'Phone', 'Zip'):
            if k in r and pd.notna(r[k]):
                r[k] = str(r[k]).strip()
            elif k in r:
                r[k] = ''
    print(f"Medi-Cal source rows (from {os.path.basename(MEDICAL_XLSX)}): {len(med_rows)}")
    print(f"D-SNP extracted rows: {len(dsnp_rows)}")

    all_rows = med_rows + dsnp_rows

    merged = {}
    order = []
    dup_count = 0
    for row in all_rows:
        key = dedup_key_npi(row)
        if key is None:
            key = ('NOKEY',) + dedup_key_fallback(row)
        if key in merged:
            dup_count += 1
            existing = merged[key]
            src_existing = existing.get('Directory Source', '')
            src_new = row.get('Directory Source', '')
            sources = [s.strip() for s in src_existing.split(';') if s.strip()]
            if src_new and src_new not in sources:
                sources.append(src_new)
            existing['Directory Source'] = '; '.join(sources)
            # fill in any blank fields on the existing row from the new row
            for k, v in row.items():
                if k == 'Directory Source':
                    continue
                if (existing.get(k) in (None, '', 'nan') or (isinstance(existing.get(k), float) and pd.isna(existing.get(k)))) and v not in (None, ''):
                    existing[k] = v
        else:
            merged[key] = dict(row)
            order.append(key)

    final_rows = [merged[k] for k in order]
    print(f"Duplicate matches merged across sources: {dup_count}")
    print(f"Final merged/deduped row count: {len(final_rows)}")
    return final_rows


def write_excel(rows):
    df = pd.DataFrame(rows)

    preferred_cols = [
        'Provider Type', 'Directory Source', 'County', 'Name', 'Credential', 'Gender',
        'Specialty', 'Office/Facility Name', 'Address', 'Street', 'City', 'State', 'Zip',
        'Phone', 'NPI', 'License', 'Board Certified', 'Office Hours', 'Languages', 'Access',
        'Hospital', 'Provider ID', 'Accepting New Patients', 'Age Group', 'Telehealth',
        'Source Page (0-indexed)', 'Source Column',
    ]
    cols = [c for c in preferred_cols if c in df.columns]
    remaining = [c for c in df.columns if c not in cols]
    df = df[cols + remaining]

    sort_cols = [c for c in ['County', 'City', 'Specialty', 'Name'] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, na_position='last').reset_index(drop=True)

    text_cols = ['Phone', 'NPI', 'Zip', 'License']
    with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Primary Care Providers')
        ws = writer.sheets['Primary Care Providers']
        header = [cell.value for cell in ws[1]]
        for col_name in text_cols:
            if col_name not in header:
                continue
            col_idx = header.index(col_name) + 1
            col_letter = get_column_letter(col_idx)
            for row in range(2, ws.max_row + 1):
                cell = ws[f"{col_letter}{row}"]
                cell.number_format = '@'
                if cell.value is not None:
                    cell.value = str(cell.value)
        for i, col_name in enumerate(header, start=1):
            col_letter = get_column_letter(i)
            maxlen = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).tolist()[:2000]])
            ws.column_dimensions[col_letter].width = min(max(maxlen + 2, 10), 45)

    print(f"Wrote {OUTPUT_XLSX} rows: {len(df)}")


def main():
    dsnp_rows = extract_all_dsnp()
    final_rows = merge_and_dedupe(dsnp_rows)
    write_excel(final_rows)


if __name__ == '__main__':
    main()
