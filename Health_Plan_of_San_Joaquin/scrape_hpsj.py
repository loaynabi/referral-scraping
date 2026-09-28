"""
HPSJ/MVHP Medi-Cal Provider Directory extractor.
Extracts Primary Care provider rows from the 2-column PDF layout using
extract_words() + x0-based column splitting, with label-anchored field parsing.

Resumable via checkpoint.json (flat list of row dicts + last_page_processed).
"""
import pdfplumber
import re
import json
import os
import sys
import time

PDF_PATH = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf\HPSJ-MVHP-Medi-Cal-Provider-Directory-September-2026.pdf"
CHECKPOINT_PATH = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\checkpoint.json"
OUTPUT_XLSX = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\Health_Plan_of_San_Joaquin_Providers.xlsx"

# 0-indexed page range for the Primary Care section (inclusive start, exclusive end).
# Determined by boundary scan; see checkpoint / report for confidence notes.
START_PAGE = 32   # 0-indexed; page 32 begins with the "Primary Care Doctors" section banner image
END_PAGE = 378    # 0-indexed, exclusive; page 378 begins the "Obstetrics/Gynecologists Doctors" section banner image

COLUMN_SPLIT_X = 300  # x0 threshold between left/right column

LABELS = [
    'PHONE:', 'BOARD CERTIFIED:', 'OFFICE HOURS:', 'LANGUAGES:',
    'HOSPITAL:', 'WEB SITE:', 'NPI:', 'LICENSE:', 'PROVIDER ID:',
    'ACCESS:', 'ACCEPTING NEW PATIENTS:', 'AGE GROUP:', 'TELEHEALTH:',
    'FAX:',
]
LABEL_PATTERN = re.compile('(' + '|'.join(re.escape(l) for l in sorted(LABELS, key=len, reverse=True)) + ')')

GENDER_RE = re.compile(r'^(.*?)\s+(Male|Female|Non-Binary|Unknown|Other)\s*$')
CITYSTATEZIP_RE = re.compile(r'^(?P<city>.+?),\s*(?P<state>[A-Z]{2})\s+(?P<zip>[\d\-]+)\s*$')

SENSITIVE_RE = re.compile(r'\bSSN\b|\bsocial security\b|\bDOB\b|\bdate of birth\b', re.I)

# Strip Unicode Private Use Area glyphs (decorative icon-font characters embedded in the
# PDF, e.g. a small arrow/bullet glyph that trails "PROVIDER ID:" values) and collapse
# whitespace.
PUA_RE = re.compile(r'[-]')


def clean_value(v):
    if not isinstance(v, str):
        return v
    v = PUA_RE.sub('', v)
    v = re.sub(r'\s+', ' ', v).strip()
    return v


def get_column_lines(page, x_min, x_max):
    """Return list of (top, text) lines for words within [x_min, x_max)."""
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    col_words = [w for w in words if x_min <= w['x0'] < x_max]
    rows = {}
    for w in col_words:
        key = round(w['top'])
        rows.setdefault(key, []).append(w)
    lines = []
    for top in sorted(rows.keys()):
        ws = sorted(rows[top], key=lambda w: w['x0'])
        text = " ".join(w['text'] for w in ws)
        lines.append((top, text))
    # Merge lines that are very close together (within 3pt) - defensive, usually not needed
    return lines


def parse_column(lines, page_num, column):
    """Parse a list of (top, text) lines from one column into provider row dicts."""
    rows = []
    texts = [t for _, t in lines]
    n = len(texts)

    # Find all name+gender line indices
    name_idxs = [i for i, t in enumerate(texts) if GENDER_RE.match(t) and 'PHONE:' not in t]

    # Find all TELEHEALTH: line indices (entry terminators)
    telehealth_idxs = [i for i, t in enumerate(texts) if t.strip().startswith('TELEHEALTH:')]

    for name_idx in name_idxs:
        # specialty header is the line before the name line
        spec_idx = name_idx - 1
        if spec_idx < 0:
            continue
        specialty = texts[spec_idx].strip()
        # sanity: specialty header should have no colon
        if ':' in specialty:
            continue

        # find the entry end = first TELEHEALTH idx >= name_idx
        end_idx = None
        for t_idx in telehealth_idxs:
            if t_idx >= name_idx:
                end_idx = t_idx
                break
        if end_idx is None:
            end_idx = n - 1  # entry got cut off at page/column end; best effort

        entry_lines = texts[spec_idx:end_idx + 1]
        name_gender_line = entry_lines[1]
        m = GENDER_RE.match(name_gender_line)
        if not m:
            continue
        name = m.group(1).strip()
        gender = m.group(2).strip()

        org = entry_lines[2].strip() if len(entry_lines) > 2 else ""

        # find first line (from idx 3 onward within entry_lines) containing a label
        first_label_i = None
        for i in range(3, len(entry_lines)):
            if LABEL_PATTERN.search(entry_lines[i]):
                first_label_i = i
                break
        if first_label_i is None:
            first_label_i = len(entry_lines)

        address_lines = [l.strip() for l in entry_lines[3:first_label_i] if l.strip()]
        address_full = ", ".join(address_lines)

        street = ""
        city = ""
        state = ""
        zipc = ""
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

        field_text = " ".join(entry_lines[first_label_i:])
        matches = list(LABEL_PATTERN.finditer(field_text))
        fields = {}
        for mi, mm in enumerate(matches):
            label = mm.group(1)
            start = mm.end()
            stop = matches[mi + 1].start() if mi + 1 < len(matches) else len(field_text)
            value = field_text[start:stop].strip().strip(',').strip()
            fields[label] = value

        row = {
            'Provider Type': 'Primary Care',
            'Name': name,
            'Gender': gender,
            'Specialty': specialty,
            'Office/Facility Name': org,
            'Address': address_full,
            'Street': street,
            'City': city,
            'State': state,
            'Zip': zipc,
            'Phone': fields.get('PHONE:', ''),
            'Board Certified': fields.get('BOARD CERTIFIED:', ''),
            'Office Hours': fields.get('OFFICE HOURS:', ''),
            'Languages': fields.get('LANGUAGES:', ''),
            'Hospital': fields.get('HOSPITAL:', ''),
            'NPI': fields.get('NPI:', ''),
            'License': fields.get('LICENSE:', ''),
            'Provider ID': fields.get('PROVIDER ID:', ''),
            'Access': fields.get('ACCESS:', ''),
            'Accepting New Patients': fields.get('ACCEPTING NEW PATIENTS:', ''),
            'Age Group': fields.get('AGE GROUP:', ''),
            'Telehealth': fields.get('TELEHEALTH:', ''),
            'Source Page (0-indexed)': page_num,
            'Source Column': column,
        }

        row = {k: clean_value(v) for k, v in row.items()}

        # privacy defensive filter
        blob = json.dumps(row)
        if SENSITIVE_RE.search(blob):
            continue

        rows.append(row)
    return rows


def parse_page(page, page_num):
    width = page.width
    left_lines = get_column_lines(page, 0, COLUMN_SPLIT_X)
    right_lines = get_column_lines(page, COLUMN_SPLIT_X, width + 10)
    rows = parse_column(left_lines, page_num, 'left')
    rows += parse_column(right_lines, page_num, 'right')
    return rows


def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('rows', []), data.get('last_page_processed', START_PAGE - 1)
    return [], START_PAGE - 1


def save_checkpoint(rows, last_page):
    tmp = CHECKPOINT_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump({'rows': rows, 'last_page_processed': last_page}, f)
    os.replace(tmp, CHECKPOINT_PATH)


def main():
    end_page = END_PAGE
    if end_page is None:
        # try to read from boundary file if present
        bpath = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\boundary.json"
        if os.path.exists(bpath):
            with open(bpath) as f:
                b = json.load(f)
            end_page = b['end_page_exclusive']
        else:
            print("WARNING: END_PAGE not set and no boundary.json found; defaulting to full doc")
            end_page = 99999

    rows, last_page = load_checkpoint()
    start_from = last_page + 1
    print(f"Resuming from page {start_from} (checkpoint had {len(rows)} rows, last_page_processed={last_page})")

    with pdfplumber.open(PDF_PATH) as pdf:
        n_pages = len(pdf.pages)
        end_page_eff = min(end_page, n_pages)
        t0 = time.time()
        for i in range(start_from, end_page_eff):
            page = pdf.pages[i]
            try:
                page_rows = parse_page(page, i)
            except Exception as e:
                print(f"ERROR on page {i}: {e}")
                page_rows = []
            rows.extend(page_rows)
            if (i - start_from + 1) % 100 == 0 or i == end_page_eff - 1:
                save_checkpoint(rows, i)
                elapsed = time.time() - t0
                print(f"Checkpointed at page {i} ({len(rows)} rows so far, {elapsed:.1f}s elapsed)")
        save_checkpoint(rows, end_page_eff - 1)
    print(f"Done. Total rows: {len(rows)}. Last page processed: {end_page_eff - 1}")


if __name__ == '__main__':
    main()
