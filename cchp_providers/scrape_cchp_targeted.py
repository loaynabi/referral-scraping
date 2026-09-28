"""
CCHP (Contra Costa Health Plan) targeted-specialty provider directory extractor.

Extracts provider rows from ONLY the 15 target specialty sections (Family Medicine,
Internal Medicine, Pediatrics, Bariatric Surgery, Orthopedics-General, Sports Medicine,
Spine, Geriatric Medicine, Occupational Medicine, Maternal Fetal Medicine, Sleep Medicine,
Rheumatology, Pain Medicine, Infectious Disease, Neurology) across 3 CCHP directory PDFs:
  - CCHP MediCal Provider Directory 20260814.pdf   (Medi-Cal, 1110 pages)
  - CCHP Provider Directory.pdf                     (Commercial A/A COB/A2T/A2P/ARCC/IHSS, 159 pages)
  - PD3RMCCPN.pdf                                   (Commercial B/B COB/Superior Court, 236 pages)

Method: each PDF's Table of Contents gives printed page numbers per specialty section
(verified empirically against actual page footers -> a constant offset between printed
page number and pdfplumber's 0-indexed page array, per source PDF). For each target
section we build a column-banded (3-column) text stream over a page window bounded by
the section's own TOC page and the immediately-following section's TOC page, then locate
the section's own header line and the next section's header line as anchors within that
stream (handles sections that start/end mid-page). Provider entries are then parsed out
of the bounded slice using label-anchored, order-independent field extraction (a provider
entry is delimited by a "Last, First[, MI], CREDENTIAL" name line).

Checkpointed per (source, section-instance) to checkpoint_raw.json so re-runs resume.
"""
import pdfplumber
import re
import json
import os
import sys
import shutil
import time
import pandas as pd
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE = r"C:\Users\DamonHollister\Desktop\referral scraping\cchp_providers"
OUTPUT_DIR = r"C:\Users\DamonHollister\Desktop\referral scraping\output"
CHECKPOINT_PATH = os.path.join(BASE, "checkpoint_raw.json")
OUTPUT_XLSX = os.path.join(BASE, "CCHP_Referral_Providers_Targeted_Specialties.xlsx")

PDF_PATHS = {
    "medical": r"C:\Users\DamonHollister\Downloads\CCHP MediCal Provider Directory 20260814.pdf",
    "commercial_a": r"C:\Users\DamonHollister\Downloads\CCHP Provider Directory.pdf",
    "commercial_b": r"C:\Users\DamonHollister\Downloads\PD3RMCCPN.pdf",
}

DIRECTORY_LABEL = {
    "medical": "Medi-Cal",
    "commercial_a": "Commercial A/A COB/A2T/A2P/ARCC/IHSS",
    "commercial_b": "Commercial B/B COB/Superior Court",
}

SOURCE_FILE_LABEL = {
    "medical": "CCHP MediCal Provider Directory 20260814.pdf",
    "commercial_a": "CCHP Provider Directory.pdf",
    "commercial_b": "PD3RMCCPN.pdf",
}

# printed_page + OFFSET = pdfplumber 0-indexed page array position (empirically verified
# against actual page-footer printed numbers for each source PDF).
OFFSET = {
    "medical": 43,
    "commercial_a": 20,
    "commercial_b": 21,
}

COL_BOUNDS = [(0, 210), (210, 390), (390, 999)]
TOP_CUTOFF_MIN = 95   # skip running page-title banner lines above this
TOP_CUTOFF_MAX = 655  # skip footer disclaimer / page-number lines at/below this

FOOTER_KEYWORDS = [
    "this list of providers", "cchp does not guarantee", "esta lista de proveedores",
    "cchp no garantiza", "disclaimer:", "descargo de responsabilidad",
    "assessment of facility accessibility", "evaluaci", "此提供者名單", "免責聲明",
    "member services at", "servicios a miembros", "member services representative",
]

# canonical specialty name -> raw header text(s) seen in the PDFs that map to it
CANONICAL_MAP = {
    "Family Medicine": "Family Medicine",
    "Internal Medicine": "Internal Medicine",
    "Pediatrics": "Pediatrics",
    "Bariatric Surgery": "Surgery - Bariatric",
    "Orthopedics - General": "Orthopaedics",              # + also "Surgery - Orthopaedic"
    "Sports Medicine": "Sports Medicine",
    "Spine": "Spine",
    "Geriatric Medicine": "Geriatric Medicine",
    "Occupational Medicine": "Occupational Medicine",
    "Maternal Fetal Medicine": "Perinatology",
    "Sleep Medicine": "Sleep Medicine",
    "Rheumatology": "Rheumatology",
    "Pain Medicine": "Pain Medicine",
    "Infectious Disease": "Infectious Disease",
    "Neurology": "Neurology",
}
RAW_TO_CANONICAL = {
    "Family Medicine": "Family Medicine",
    "Internal Medicine": "Internal Medicine",
    "Pediatrics": "Pediatrics",
    "Surgery - Bariatric": "Bariatric Surgery",
    "Orthopaedics": "Orthopedics - General",
    "Surgery - Orthopaedic": "Orthopedics - General",
    "Sports Medicine": "Sports Medicine",
    "Geriatric Medicine": "Geriatric Medicine",
    "Perinatology": "Maternal Fetal Medicine",
    "Sleep Medicine": "Sleep Medicine",
    "Rheumatology": "Rheumatology",
    "Pain Medicine": "Pain Medicine",
    "Infectious Disease": "Infectious Disease",
    "Neurology": "Neurology",
}

# Each tuple: (raw_header, network, start_printed_page, next_printed_page, next_raw_header, provider_type)
TARGET_SECTIONS = {
    "medical": [
        ("Family Medicine", "CCRMC", 1, 26, "Internal Medicine", "Primary Care"),
        ("Internal Medicine", "CCRMC", 26, 27, "Pediatrics", "Primary Care"),
        ("Pediatrics", "CCRMC", 27, 32, "Family Medicine", "Primary Care"),
        ("Family Medicine", "CPN", 32, 64, "Geriatric Medicine", "Primary Care"),
        ("Geriatric Medicine", "CPN", 64, 64, "Internal Medicine", "Primary Care"),
        ("Internal Medicine", "CPN", 64, 81, "Pediatrics", "Primary Care"),
        ("Pediatrics", "CPN", 81, 116, "Cardiology", "Primary Care"),
        ("Neurology", "CCRMC", 129, 131, "Obstetrics and Gynecology", "Referral"),
        ("Orthopaedics", "CCRMC", 140, 142, "Otolaryngology (Ear, Nose & Throat)", "Referral"),
        ("Rheumatology", "CCRMC", 148, 150, "Sleep Medicine", "Referral"),
        ("Sleep Medicine", "CCRMC", 150, 151, "Surgery - General", "Referral"),
        ("Surgery - Orthopaedic", "CCRMC", 154, 156, "Surgery - Plastic", "Referral"),
        ("Infectious Disease", "CPN", 299, 309, "Interventional Cardiology", "Referral"),
        ("Neurology", "CPN", 352, 356, "Obstetrics and Gynecology", "Referral"),
        ("Orthopaedics", "CPN", 408, 419, "Otolaryngology (Ear, Nose & Throat)", "Referral"),
        ("Pain Medicine", "CPN", 435, 438, "Pediatric Allergy & Immunology", "Referral"),
        ("Perinatology", "CPN", 468, 468, "Phlebology", "Referral"),
        ("Rheumatology", "CPN", 519, 530, "Sleep Medicine", "Referral"),
        ("Sleep Medicine", "CPN", 530, 538, "Sports Medicine", "Referral"),
        ("Sports Medicine", "CPN", 538, 543, "Surgery - Bariatric", "Referral"),
        ("Surgery - Bariatric", "CPN", 543, 545, "Surgery - Cardiothoracic", "Referral"),
        ("Surgery - Orthopaedic", "CPN", 582, 613, "Surgery - Plastic", "Referral"),
    ],
    "commercial_a": [
        ("Family Medicine", "CCRMC", 1, 27, "Internal Medicine", "Primary Care"),
        ("Internal Medicine", "CCRMC", 27, 28, "Pediatrics", "Primary Care"),
        ("Pediatrics", "CCRMC", 28, 35, "Dentist", "Primary Care"),
    ],
    "commercial_b": [
        ("Family Medicine", "CCRMC", 1, 12, "Internal Medicine", "Primary Care"),
        ("Internal Medicine", "CCRMC", 12, 12, "Pediatrics", "Primary Care"),
        ("Pediatrics", "CCRMC", 12, 14, "Family Medicine", "Primary Care"),
        ("Family Medicine", "CPN", 14, 44, "Geriatric Medicine", "Primary Care"),
        ("Geriatric Medicine", "CPN", 44, 45, "Internal Medicine", "Primary Care"),
        ("Internal Medicine", "CPN", 45, 61, "Pediatrics", "Primary Care"),
        ("Pediatrics", "CPN", 61, 92, "Acupuncture", "Primary Care"),
        ("Perinatology", "CPN", 123, 124, "Optometry", "Referral"),
    ],
}

# Specialties confirmed ABSENT from a source's TOC (for the final report)
MISSING_SECTIONS = {
    "medical": ["Spine", "Occupational Medicine"],
    "commercial_a": [
        "Bariatric Surgery", "Orthopedics - General", "Sports Medicine", "Spine",
        "Geriatric Medicine", "Occupational Medicine", "Maternal Fetal Medicine",
        "Sleep Medicine", "Rheumatology", "Pain Medicine", "Infectious Disease", "Neurology",
    ],
    "commercial_b": [
        "Bariatric Surgery", "Orthopedics - General", "Sports Medicine", "Spine",
        "Occupational Medicine", "Sleep Medicine", "Rheumatology", "Pain Medicine",
        "Infectious Disease", "Neurology",
    ],
}

# ---------------------------------------------------------------------------
# Regexes
# ---------------------------------------------------------------------------
NAME_RE = re.compile(
    r"^([A-Z][A-Za-z'\-\.]*(?:\s[A-Z][A-Za-z'\-\.]*){0,3}),\s"
    r"([A-Z][A-Za-z'\-\. ]{0,30}),\s"
    r"([A-Z][A-Za-z\.\-]{0,10})$"
)
PHONE_RE = re.compile(r"^\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]\d{4}$")
CITYSTATEZIP_RE = re.compile(r"^(?P<city>[A-Za-z][A-Za-z\.\-' ]*),\s*(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)$")
STREET_RE = re.compile(r"^\d+[\w\s.,#\-/]*$")
WEBSITE_RE = re.compile(r"^(?:Website:\s*)?(?:https?://|www\.)\S+$", re.I)

LABEL_PATTERNS = [
    ("Provider ID", re.compile(r"^Provider ID:\s*(.*)$", re.I)),
    ("NPI", re.compile(r"^NPI:\s*(.*)$", re.I)),
    ("CA Lic", re.compile(r"^CA Lic(?:ense)?:?\s*(.*)$", re.I)),
    ("Provider Languages", re.compile(r"^(?:Provider )?Languages:\s*(.*)$", re.I)),
    ("Race", re.compile(r"^Race:\s*(.*)$", re.I)),
    ("Ethnicity", re.compile(r"^Ethnicity:\s*(.*)$", re.I)),
    ("Gender Affirming Care", re.compile(r"^Gender Affirming Care:\s*(.*)$", re.I)),
    ("Gender", re.compile(r"^Gender:\s*(.*)$", re.I)),
    ("Accepting New Patients", re.compile(r"^Accepting New Patients:\s*(.*)$", re.I)),
    ("CDE", re.compile(r"^CDE:?\s*(.*)$", re.I)),
    ("DEI", re.compile(r"^DEI(?: Training)?:?\s*(.*)$", re.I)),
    ("Accessibility", re.compile(r"^Accessibility:\s*(.*)$", re.I)),
]
FLAG_LINES = {
    "board certified": ("Board Certified", "Board Certified"),
    "not boarded specialty": ("Board Certified", "Not Boarded Specialty"),
    "not ccs paneled": ("CCS Status", "Not CCS Paneled"),
    "ccs paneled": ("CCS Status", "CCS Paneled"),
    "enrolled in ccs": ("CCS Status", "Enrolled in CCS"),
    "not css paneled": ("CCS Status", "Not CCS Paneled"),
}
HOURS_LABEL_RE = re.compile(r"^Hours:\s*(.*)$", re.I)


def normalize_header(s):
    s = re.sub(r"[\u2010-\u2015]", "-", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def clean_value(v):
    if not isinstance(v, str):
        return v
    v = re.sub(r"\s+", " ", v).strip()
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
        if top < TOP_CUTOFF_MIN or top >= TOP_CUTOFF_MAX:
            continue
        ws = sorted(rows[top], key=lambda w: w['x0'])
        text = " ".join(w['text'] for w in ws).strip()
        if not text:
            continue
        low = text.lower()
        if any(kw in low for kw in FOOTER_KEYWORDS):
            continue
        lines.append(text)
    return lines


HYPHEN_WRAP_RE = re.compile(r"^(.*[A-Za-z])-$")


def merge_hyphen_wraps(lines):
    """PDF word-wrap sometimes breaks a word (esp. surnames/ethnicity terms) across two
    stacked lines with a trailing hyphen, e.g. 'Orengo-McFarlane, Mi-' / 'chelle, MD'.
    Merge such pairs (dehyphenate) so downstream Name/label matching sees one line."""
    out = []
    i = 0
    n = len(lines)
    while i < n:
        cur = lines[i]
        m = HYPHEN_WRAP_RE.match(cur)
        if m and i + 1 < n:
            nxt = lines[i + 1]
            if nxt and nxt[0].islower():
                out.append(cur[:-1] + nxt)
                i += 2
                continue
        out.append(cur)
        i += 1
    return out


WEBSITE_BARE_RE = re.compile(r"^Website:\s*$", re.I)
URLFRAG_RE = re.compile(r"^[\w\.\-/]+$")
URL_END_RE = re.compile(r"\.[A-Za-z]{2,10}$")


def merge_website_wraps(lines):
    """A bare 'Website:' label is sometimes followed by a domain that itself wraps
    across 1-3 more lines (e.g. 'Website:' / 'www.example.' / 'com'). Merge those into
    a single 'Website: www.example.com' line."""
    out = []
    i = 0
    n = len(lines)
    while i < n:
        cur = lines[i]
        if WEBSITE_BARE_RE.match(cur):
            j = i + 1
            frag = []
            while j < n and len(frag) < 4 and URLFRAG_RE.match(lines[j]):
                frag.append(lines[j])
                j += 1
                if URL_END_RE.search("".join(frag)):
                    break
            if frag:
                out.append("Website: " + "".join(frag))
                i = j
                continue
        out.append(cur)
        i += 1
    return out


def merge_trailing_comma_wraps(lines):
    """A comma-separated list value (Ethnicity/Race/Provider Languages) that wraps to a
    second line leaves the first line ending in a bare ',' -- merge with the next line
    so the continuation doesn't get mistaken for a new Name/label line (its 2-comma
    'X, Y, Z' shape can otherwise false-match the provider Name pattern)."""
    out = []
    i = 0
    n = len(lines)
    while i < n:
        cur = lines[i]
        while cur.rstrip().endswith(",") and i + 1 < n:
            # a line ending in a bare trailing comma is (in this document format)
            # always a wrapped comma-separated list continuing onto the next line --
            # merge unconditionally (the continuation can otherwise coincidentally
            # match the provider-Name pattern, e.g. "German, Norwegian, Swedish").
            nxt = lines[i + 1]
            cur = cur.rstrip() + " " + nxt
            i += 1
        out.append(cur)
        i += 1
    return out


def build_stream(pdf, page_indices):
    """Flat ordered list of text lines across the given pdfplumber page indices,
    column 0 -> 1 -> 2 per page, pages in given order."""
    stream = []
    for pidx in page_indices:
        if pidx < 0 or pidx >= len(pdf.pages):
            continue
        page = pdf.pages[pidx]
        for (xmin, xmax) in COL_BOUNDS:
            col_lines = get_column_lines(page, xmin, xmax)
            col_lines = merge_hyphen_wraps(col_lines)
            col_lines = merge_website_wraps(col_lines)
            col_lines = merge_trailing_comma_wraps(col_lines)
            stream.extend(col_lines)
    return stream


def find_header(stream, target_name, start_idx=0):
    """Find the position AFTER the header line(s) matching target_name (allowing a
    1-3 line wrap), searching from start_idx forward. Returns None if not found."""
    target_norm = normalize_header(target_name)
    n = len(stream)
    for i in range(start_idx, n):
        for span in (1, 2, 3):
            if i + span > n:
                continue
            joined = normalize_header(" ".join(stream[i:i + span]))
            if joined == target_norm:
                return i + span
    return None


def find_header_flex(stream, target_name, start_idx=0):
    pos = find_header(stream, target_name, start_idx)
    if pos is not None:
        return pos
    # fallback: substring match (defensive, in case of odd wrapping/punctuation)
    target_norm = normalize_header(target_name)
    for i in range(start_idx, len(stream)):
        if normalize_header(stream[i]) == target_norm:
            return i + 1
    return None


def classify_line(line):
    """Return (kind, key, value) for a recognized labeled/flag/structural line, else None."""
    for key, pat in LABEL_PATTERNS:
        m = pat.match(line)
        if m:
            return ("label", key, m.group(1).strip())
    low = line.strip().lower()
    if low in FLAG_LINES:
        key, val = FLAG_LINES[low]
        return ("flag", key, val)
    if HOURS_LABEL_RE.match(line):
        inline = HOURS_LABEL_RE.match(line).group(1).strip()
        return ("hours_start", "Hours", inline)
    if PHONE_RE.match(line):
        return ("phone", "Phone", line.strip())
    if CITYSTATEZIP_RE.match(line):
        m = CITYSTATEZIP_RE.match(line)
        return ("citystatezip", "CityStateZip", (m.group('city').strip(), m.group('state').strip(), m.group('zip').strip()))
    if WEBSITE_RE.match(line):
        val = re.sub(r"^Website:\s*", "", line.strip(), flags=re.I)
        return ("website", "Website", val)
    return None


def parse_entry_block(block):
    """block[0] is the Name line (already validated by NAME_RE); block[1:] are the
    remaining lines for this entry, in stream order."""
    m = NAME_RE.match(block[0])
    last, first, credential = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
    name = f"{last}, {first}"

    fields = {
        "Provider ID": "", "NPI": "", "CA Lic": "", "Provider Languages": "",
        "Race": "", "Ethnicity": "", "Gender Affirming Care": "", "Gender": "",
        "Accepting New Patients": "", "CDE": "", "DEI": "", "Accessibility": "",
        "Board Certified": "", "CCS Status": "", "Phone": "", "Website": "",
    }
    hours_parts = []
    facility_lines = []
    street_lines = []
    city, state, zipc = "", "", ""
    board_affil_lines = []

    PHASE_FACILITY, PHASE_ADDRESS, PHASE_HOURS, PHASE_BOARD, PHASE_DONE = range(5)
    phase = PHASE_FACILITY
    seen_provider_id = False
    seen_ca_lic_or_lang = False
    seen_race = False
    pending_accessibility_label = False

    for line in block[1:]:
        line = line.strip()
        if not line:
            continue

        if pending_accessibility_label:
            # previous line was a bare "Accessibility:" with no inline codes
            fields["Accessibility"] = line
            pending_accessibility_label = False
            continue

        c = classify_line(line)
        if c is not None:
            kind, key, val = c
            if phase == PHASE_FACILITY:
                # any recognized labeled/flag/phone/citystatezip line means the facility
                # name (if any) is already fully collected
                phase = PHASE_ADDRESS
            if kind == "label":
                if key == "Provider ID":
                    seen_provider_id = True
                    phase = PHASE_ADDRESS
                elif key in ("CA Lic", "Provider Languages"):
                    seen_ca_lic_or_lang = True
                    phase = PHASE_BOARD
                elif key == "Race":
                    seen_race = True
                    phase = PHASE_DONE
                if key == "Accessibility" and val == "":
                    pending_accessibility_label = True
                    continue
                fields[key] = val
            elif kind == "flag":
                fields[key] = val
            elif kind == "hours_start":
                phase = PHASE_HOURS
                if val:
                    hours_parts.append(val)
            elif kind == "phone":
                fields["Phone"] = val
            elif kind == "citystatezip":
                city, state, zipc = val
            elif kind == "website":
                fields["Website"] = val
            continue

        # unlabeled line -> bucket by current phase
        if phase == PHASE_FACILITY:
            if STREET_RE.match(line) and len(facility_lines) >= 1:
                # looks like the start of a street address (leading digit) -- treat as
                # address, not a further facility-name wrap
                street_lines.append(line)
                phase = PHASE_ADDRESS
            elif len(facility_lines) < 3:
                facility_lines.append(line)
            else:
                # facility name already took 3 lines; anything more is address
                street_lines.append(line)
                phase = PHASE_ADDRESS
        elif phase == PHASE_ADDRESS:
            street_lines.append(line)
        elif phase == PHASE_HOURS:
            hours_parts.append(line)
        elif phase == PHASE_BOARD and not seen_race:
            board_affil_lines.append(line)
        # else: stray/unrecognized line after Race -> ignore

    row = {
        "Name": name,
        "Credentials": credential,
        "Facility/Practice Name": clean_value(" ".join(facility_lines)),
        "Street": clean_value(" ".join(street_lines)),
        "City": city,
        "State": state,
        "Zip": zipc,
        "Phone": fields["Phone"],
        "Website": fields["Website"],
        "Hours": clean_value("; ".join(hours_parts)),
        "Accessibility": fields["Accessibility"],
        "NPI": fields["NPI"],
        "CA License": fields["CA Lic"],
        "Board/Certification Affiliation": clean_value(" ".join(board_affil_lines)),
        "Provider Languages": fields["Provider Languages"],
        "Race": fields["Race"],
        "Ethnicity": fields["Ethnicity"],
        "Gender": fields["Gender"],
        "Gender Affirming Care": fields["Gender Affirming Care"],
        "Board Certified": fields["Board Certified"],
        "Accepting New Patients": fields["Accepting New Patients"],
        "CCS Status": fields["CCS Status"],
        "DEI Training": fields["DEI"],
        "CDE": fields["CDE"],
        "Provider ID": fields["Provider ID"],
    }
    return {k: clean_value(v) for k, v in row.items()}


def parse_entries(stream_slice):
    n = len(stream_slice)
    starts = [i for i in range(n) if NAME_RE.match(stream_slice[i])]
    rows = []
    for idx, s in enumerate(starts):
        e = starts[idx + 1] if idx + 1 < len(starts) else n
        block = stream_slice[s:e]
        try:
            row = parse_entry_block(block)
            rows.append(row)
        except Exception as ex:
            print(f"    ! failed to parse entry block starting '{block[0][:40]}': {ex}")
    return rows


def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"rows": [], "completed": []}


def save_checkpoint(data):
    tmp = CHECKPOINT_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f)
    os.replace(tmp, CHECKPOINT_PATH)


def process_source(source_key, data):
    path = PDF_PATHS[source_key]
    offset = OFFSET[source_key]
    sections = TARGET_SECTIONS[source_key]
    rows = data["rows"]
    completed = set(data["completed"])

    with pdfplumber.open(path) as pdf:
        for sec_i, (raw_header, network, start_page, next_page, next_raw, ptype) in enumerate(sections):
            inst_key = f"{source_key}::{sec_i}::{raw_header}::{network}::{start_page}"
            if inst_key in completed:
                print(f"Skipping {inst_key} (checkpointed)")
                continue

            start_idx = start_page + offset
            next_idx = next_page + offset
            win_start = max(0, start_idx - 2)
            win_end = min(len(pdf.pages) - 1, next_idx + 2)
            page_indices = list(range(win_start, win_end + 1))

            t0 = time.time()
            stream = build_stream(pdf, page_indices)

            own_pos_hint = max(0, (start_idx - win_start) * 3 * 12)  # rough hint, unused for correctness
            content_start = find_header_flex(stream, raw_header, 0)
            if content_start is None:
                print(f"  ! Could not locate header '{raw_header}' for {inst_key} in window pages {page_indices[0]}-{page_indices[-1]}")
                completed.add(inst_key)
                save_checkpoint({"rows": rows, "completed": sorted(completed)})
                continue

            content_end = find_header_flex(stream, next_raw, content_start)
            if content_end is None:
                content_end = len(stream)
                print(f"  ! Could not locate next-header '{next_raw}' after {inst_key}; using window end")

            slice_ = stream[content_start:content_end]
            entries = parse_entries(slice_)

            canonical = RAW_TO_CANONICAL[raw_header]
            for r in entries:
                r["Source File"] = SOURCE_FILE_LABEL[source_key]
                r["Directory/Plan"] = DIRECTORY_LABEL[source_key]
                r["Network"] = network
                r["Provider Type"] = ptype
                r["Specialty"] = canonical
                r["_Raw Section Header"] = raw_header

            rows.extend(entries)
            completed.add(inst_key)
            save_checkpoint({"rows": rows, "completed": sorted(completed)})
            elapsed = time.time() - t0
            print(f"  {inst_key}: {len(entries)} rows ({elapsed:.1f}s, pages {page_indices[0]}-{page_indices[-1]})")

    return rows


def dedupe_within_source(rows):
    """Collapse exact-duplicate rows (all fields identical) within the same source file
    only -- a parsing artifact safeguard. Never dedupes across sources/specialties."""
    seen = {}
    out = []
    dup_count = 0
    for r in rows:
        key = (r.get("Source File"), json.dumps(r, sort_keys=True))
        if key in seen:
            dup_count += 1
            continue
        seen[key] = True
        out.append(r)
    print(f"Exact-duplicate rows collapsed (within-source only): {dup_count}")
    return out


def write_excel(rows):
    df = pd.DataFrame(rows)
    if "_Raw Section Header" in df.columns:
        df = df.drop(columns=["_Raw Section Header"])

    preferred_cols = [
        "Source File", "Directory/Plan", "Network", "Provider Type", "Specialty",
        "Name", "Credentials", "Gender", "Facility/Practice Name",
        "Street", "City", "State", "Zip", "Phone", "Fax", "Website",
        "NPI", "CA License", "Provider ID",
        "Board/Certification Affiliation", "Board Certified",
        "Accepting New Patients", "Hours", "Provider Languages",
        "Race", "Ethnicity", "Gender Affirming Care", "CCS Status",
        "DEI Training", "CDE", "Accessibility",
    ]
    cols = [c for c in preferred_cols if c in df.columns]
    remaining = [c for c in df.columns if c not in cols]
    df = df[cols + remaining]

    sort_cols = [c for c in ["Directory/Plan", "Specialty", "Network", "Name"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, na_position="last").reset_index(drop=True)

    text_cols = ["Phone", "Fax", "NPI", "Zip"]
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Targeted Specialty Providers")
        ws = writer.sheets["Targeted Specialty Providers"]
        header = [cell.value for cell in ws[1]]
        for col_name in text_cols:
            if col_name not in header:
                continue
            col_idx = header.index(col_name) + 1
            col_letter = get_column_letter(col_idx)
            for row in range(2, ws.max_row + 1):
                cell = ws[f"{col_letter}{row}"]
                cell.number_format = "@"
                if cell.value is not None:
                    cell.value = str(cell.value)
        for i, col_name in enumerate(header, start=1):
            col_letter = get_column_letter(i)
            maxlen = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str).tolist()[:3000]])
            ws.column_dimensions[col_letter].width = min(max(maxlen + 2, 10), 45)

    print(f"Wrote {OUTPUT_XLSX} rows: {len(df)}")
    return df


def main():
    data = load_checkpoint()
    for source_key in ("medical", "commercial_a", "commercial_b"):
        print(f"\n=== Processing source: {source_key} ===")
        process_source(source_key, data)

    rows = data["rows"]
    print(f"\nTotal raw rows across all sources: {len(rows)}")

    by_source = {}
    for r in rows:
        by_source.setdefault(r.get("Source File"), []).append(r)
    deduped_rows = []
    for src, src_rows in by_source.items():
        deduped_rows.extend(dedupe_within_source(src_rows))

    df = write_excel(deduped_rows)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    dest = os.path.join(OUTPUT_DIR, os.path.basename(OUTPUT_XLSX))
    shutil.copyfile(OUTPUT_XLSX, dest)
    print(f"Copied to {dest}")

    print("\n=== Row counts by Source File ===")
    print(df.groupby("Source File").size())
    print("\n=== Row counts by Specialty ===")
    print(df.groupby("Specialty").size())
    print("\n=== Row counts by Source File x Specialty ===")
    print(df.groupby(["Source File", "Specialty"]).size())

    print("\n=== Sections confirmed absent from source TOC (skipped) ===")
    for src, missing in MISSING_SECTIONS.items():
        print(f"{src}: {', '.join(missing)}")


if __name__ == "__main__":
    main()
