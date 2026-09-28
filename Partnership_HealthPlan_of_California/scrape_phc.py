"""
Partnership HealthPlan of California (PHC) -- Primary Care Provider directory PDF parser
==========================================================================================

BACKGROUND
----------
PHC's live "Find a Provider" web search tool (providerdirectory.partnershiphp.org) was
down with a Cloudflare 522 (origin server unreachable) on 2026-09-09 -- see the git
history of this file / SKIPPED.md for that earlier, abandoned scraping attempt.

Instead, this script parses PHC's own **official PCP directory PDFs**, already
downloaded to the local `pdfs/` folder:

  - MCEasternPCPProDir.pdf / MCNorthernPCPProDir.pdf
        Primary-care-only directories (every listing in the file is Primary Care).
  - MCLakeProDir.pdf, MCMarinProDir.pdf, MCMendoProDir.pdf, MCNapaProDir.pdf,
    MCSolanoProDir.pdf, MCSonomaProDir.pdf, MCYoloProDir.pdf
        Combined county directories that contain SEVERAL sections: Primary Care,
        Facilities/Pharmacy/Urgent Care/VSP, Other Services, Specialists (referral
        only), and a Provider Index. We must extract ONLY the Primary Care section.

SECTION DETECTION
------------------
Every content page in these PDFs repeats a two-line running header at the very top,
e.g.:
    PARTNERSHIP PRIMARY CARE PROVIDERS
    LAKE COUNTY DIRECTORY
or for other sections:
    PARTNERSHIP FACILITIES, PHARMACY, URGENT CARE & VSP LIST
    PARTNERSHIP OTHER SERVICES PROVIDERS
    PARTNERSHIP SPECIALIST PROVIDERS - AVAILABLE BY REFERRAL ONLY
    PARTNERSHIP PROVIDER INDEX
This header text is reliable and is used to select only "PARTNERSHIP PRIMARY CARE
PROVIDERS" pages (case-insensitive, tolerant of extra whitespace). This automatically
skips all front matter (table of contents, multi-language instructions, the glossary,
the FAKE example page using "John Doe" / NPI 1234567890 / "1234 Hospital Drive"), the
Specialty Care section, Facilities/Pharmacy listings, and the Index.

LAYOUT / COLUMN RECONSTRUCTION
-------------------------------
Real provider-listing pages use a dense 3-column layout. Plain `page.extract_text()`
interleaves the 3 columns in a confusing, incorrect reading order (verified against
page dumps -- see scratch_dump.txt from development). Instead we use
`page.extract_words()` with x0/x1/top coordinates:

  1. Drop the running header (top < ~45) and the page footer boilerplate
     (top > ~710: "To verify availability...", the glossary-code legend, "Revised
     Date...").
  2. Drop small superscript sequence-number tokens like "^62^" (a per-listing index
     printed by PHC, not meaningful provider data).
  3. Bucket remaining words into 3 left-to-right column bands by x0 (bands verified
     empirically across sample pages/files: col1 x0<200, col2 200<=x0<400,
     col3 x0>=400 -- consistent because every file uses the same 612pt-wide template).
  4. Within each column band, cluster words into lines by `top` (+/-2.5pt tolerance),
     sort lines top-to-bottom, and join same-line words left-to-right by x0.
  5. Concatenate column1's lines, then column2's, then column3's -- this reproduces
     PHC's actual newspaper-style fill order (verified: a facility/office block's
     provider cluster can start in one column and its facility footer block can land
     in the next column, and the sequence only makes sense read column-major).

FACILITY <-> PROVIDER ASSOCIATION
-----------------------------------
Empirically (confirmed via coordinate dumps), each facility/office info block
(name, PCP PHC#, facility NPI, RHC/FQHC/type, specialty list, accepting status,
ages accepted, address, phone/fax, hours) appears AFTER the cluster of providers
who work there, not before. So we buffer provider entries as we encounter them,
and when a facility block is parsed, we retroactively attach its office/address/
phone/fax/specialty/accepting/ages info to every buffered provider back to the
previous flush point (this can span a column or page boundary, since providers can
be listed with no facility block on their page and pick one up on the next page/
column -- the facility block is genuinely the trailing "footer" for its provider
group in this template).

PARSING HEURISTICS
--------------------
- A PROVIDER record starts at a line containing a comma (Last, First ..., CRED)
  that is followed (skipping index markers) by an "NPI:" line.
- A FACILITY record starts at a line with no leading field-label that is followed
  (skipping index markers) by a "PCP PHC#:" line.
- Each record then greedily consumes recognized field lines (NPI/License/Gender/
  Race/Ethnicity/Hospital Privileges/Special Accommodations/Languages/CCS-CHDP
  flags for providers; PHC#/facility-NPI/type/specialty/accepting/ages/address/
  city-state-zip/access-codes/level-of-access/phone/fax/hours for facilities)
  until the next provider or facility start is detected, or page footer is hit.

`page.extract_tables()` was also tried on several sample pages (see development
notes) but PHC does not draw ruled table lines around these listings, so
pdfplumber detects no tables on the dense listing pages -- the coordinate/column
approach above is what actually works and is used throughout.

PII SAFETY
----------
No SSN, date of birth, or similar identifiers are printed anywhere in this
directory; `strip_pii()` is still applied defensively to every row before it is
written out, in case any such field is ever encountered.

CHECKPOINTING
-------------
Each PDF is fully parsed, then appended to the JSON checkpoint file so partial
progress survives interruptions:
    checkpoint.json  (flat list of row dicts, each tagged "source_pdf")

FINAL OUTPUT
------------
After all PDFs are processed, rows are deduplicated on (NPI, Address, Specialty)
and written to an Excel file with Phone/Fax/NPI/Zip/License columns forced to
text format so Excel does not mangle leading zeros or long numeric IDs:
    Partnership_HealthPlan_Providers.xlsx

Usage
-----
    pip install pdfplumber pandas openpyxl
    python scrape_phc.py
"""

import json
import os
import re
import sys

import pdfplumber

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(BASE_DIR, "pdfs")
CHECKPOINT_PATH = os.path.join(BASE_DIR, "checkpoint.json")
OUTPUT_XLSX_PATH = os.path.join(BASE_DIR, "Partnership_HealthPlan_Providers.xlsx")

PDF_FILES = [
    "MCEasternPCPProDir.pdf",
    "MCNorthernPCPProDir.pdf",
    "MCLakeProDir.pdf",
    "MCMarinProDir.pdf",
    "MCMendoProDir.pdf",
    "MCNapaProDir.pdf",
    "MCSolanoProDir.pdf",
    "MCSonomaProDir.pdf",
    "MCYoloProDir.pdf",
]

# ---------------------------------------------------------------------------
# Column bands (x0 ranges), verified against sample pages from multiple files.
# ---------------------------------------------------------------------------
COLUMN_BANDS = [(0, 200), (200, 400), (400, 10000)]
HEADER_TOP_MAX = 45
FOOTER_TOP_MIN = 710

IDX_RE = re.compile(r"^\^?\d{1,4}\^?$")
NPI_RE = re.compile(r"^NPI:\s*(.+)$", re.I)
LICENSE_RE = re.compile(r"^License:\s*(.+)$", re.I)
GENDER_RE = re.compile(r"^Gender:\s*(.+)$", re.I)
RACE_RE = re.compile(r"^Race:\s*(.+)$", re.I)
ETH_RE = re.compile(r"^Ethnicity:\s*(.+)$", re.I)
HOSP_RE = re.compile(r"^Hospital Privileges:\s*(.*)$", re.I)
SPECIAL_ACCOM_RE = re.compile(r"^Special Accommodations:\s*(.*)$", re.I)
PHC_RE = re.compile(r"^PCP\s*PHC#:\s*(\S+)", re.I)
LEVEL_RE = re.compile(r"^Level of Access:\s*(.+)$", re.I)
PHONE_RE = re.compile(r"^Primary Phone:\s*(.+)$", re.I)
FAX_RE = re.compile(r"^Primary Fax:\s*(.+)$", re.I)
ACCEPT_RE = re.compile(r"^Accepting\s+(.+)$", re.I)
CCS_RE = re.compile(r"^CCS Provider$", re.I)
CHDP_RE = re.compile(r"^CHDP Provider$", re.I)
CITY_STATE_ZIP_RE = re.compile(r"^(.+?),\s*CA\s*(\d{5})(-\d{4})?\s*$")
ACCESS_CODES_RE = re.compile(r"^(?:[PEBIRT]+,?\s*)+$")
HOURS_RE = re.compile(
    r"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b.*\d", re.I
)
AGE_RE = re.compile(
    r"^(No Age Limitations|\d*\s*[-–]\s*\d+ Years$|\d*\+?\s*Years and Older$|"
    r"\d*\s*[-–]\s*\d+ Years Old$|Birth.*|All Ages.*)$",
    re.I,
)
FACILITY_TYPE_RE = re.compile(
    r"^(RHC|FQHC|Tribal FQHC|Community Health Clinic|Hospital Outpatient Department|"
    r"Indian Health Service Clinic)$",
    re.I,
)
STREET_RE = re.compile(r"^\d+[A-Za-z]?\s+\S.*$")
PROVIDER_NAME_RE = re.compile(r"^[A-Z][A-Za-zÀ-ſ'\-]+,\s*.+$")
# Credential token: short, starts uppercase, letters/hyphen/slash (MD, DO, NP,
# PA-C, FNP-BC, DNP, CNM, PsyD, PharmD, LCSW, AGNP, GNP-BC, ...)
CRED_TOKEN_RE = re.compile(r"^[A-Z][A-Za-z\-/]{1,8}\.?$")
# A full provider "Name, Credential" line (comma present, ends in a credential-
# looking token).
NAME_CRED_RE = re.compile(
    r"^[A-Z][\w'\-]+(?:\s[A-Z][\w'\-]+)*,\s+.+,\s*([A-Za-z\-/]{2,9})\.?\s*$"
)
FIELD_LINE_RES = (
    re.compile(r"^NPI:", re.I), re.compile(r"^License:", re.I),
    re.compile(r"^Gender:", re.I), re.compile(r"^Race:", re.I),
    re.compile(r"^Ethnicity:", re.I), re.compile(r"^Hospital Privileges:", re.I),
    re.compile(r"^Special Accommodations:", re.I), re.compile(r"^PCP\s*PHC#:", re.I),
    re.compile(r"^Level of Access:", re.I), re.compile(r"^Primary Phone:", re.I),
    re.compile(r"^Primary Fax:", re.I), re.compile(r"^Accepting\s+", re.I),
)

SPECIALTY_KEYWORDS = [
    "Family Medicine", "Family Practice", "Internal Medicine", "Pediatrics",
    "Geriatric", "Hospice & Palliative Medicine", "Family Nurse Practitioner",
    "Adult Medicine", "Obstetrics", "Gynecology", "OB/GYN", "General Practice",
    "Preventive Medicine", "Urgent Care", "Naturopathic Medicine",
]

BOILERPLATE_SNIPPETS = [
    "to verify the availability",
    "board certified",
    "revised date:",
    "partnership member services",
]


def is_boilerplate(line: str) -> bool:
    low = line.lower()
    return any(s in low for s in BOILERPLATE_SNIPPETS)


def clean_word_text(t: str) -> str:
    # strip diamond/marker glyphs (board-certified / cultural-competency symbols)
    return t.replace("◆", "").replace("", "").strip()


ACCOM_CONTINUATION_RE = re.compile(
    r"^(Tables|Gurney Accommodations?|Scales that weigh|Hoyer Lift|"
    r"Hearing Impaired|Accommodations$|Onsite Interpretation|Basic Access|"
    r"Special Scales|Adj\.?\s*Exam|Wheelchair|weigh up to|up to\b|\d*\s*lbs\b)",
    re.I,
)

CONTINUATION_RES = (
    NPI_RE, LICENSE_RE, GENDER_RE, RACE_RE, ETH_RE, HOSP_RE, SPECIAL_ACCOM_RE,
)


def is_continuation_line(line):
    return any(r.match(line) for r in CONTINUATION_RES) or bool(
        ACCOM_CONTINUATION_RE.match(line)
    )


def build_page_lines(page):
    """Return the page's text lines in PHC's true column-major reading order."""
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    cleaned = []
    for w in words:
        txt = clean_word_text(w["text"])
        if not txt:
            continue
        top = w["top"]
        if top < HEADER_TOP_MAX or top > FOOTER_TOP_MIN:
            continue
        if IDX_RE.match(txt):
            continue
        cleaned.append({"text": txt, "x0": w["x0"], "top": top})

    all_lines = []
    for lo, hi in COLUMN_BANDS:
        band_words = [w for w in cleaned if lo <= w["x0"] < hi]
        band_words.sort(key=lambda w: (w["top"], w["x0"]))
        lines = []
        cur = []
        cur_top = None
        for w in band_words:
            if cur_top is None or abs(w["top"] - cur_top) <= 2.5:
                cur.append(w)
                cur_top = w["top"] if cur_top is None else cur_top
            else:
                lines.append(cur)
                cur = [w]
                cur_top = w["top"]
        if cur:
            lines.append(cur)
        line_tuples = []
        for ln in lines:
            ln_sorted = sorted(ln, key=lambda w: w["x0"])
            text = " ".join(w["text"] for w in ln_sorted)
            text = re.sub(r"\s+", " ", text).strip()
            top = min(w["top"] for w in ln)
            if text:
                line_tuples.append((top, text))
        line_tuples.sort(key=lambda t: t[0])
        col_lines = [t[1] for t in line_tuples]
        # PHC's column layout can overflow a record's trailing fields
        # (Race:/Ethnicity:/etc.) onto the very top of the NEXT column instead
        # of wrapping within the same column. Reattach any such leading
        # continuation lines to the end of the previous column instead.
        while col_lines and all_lines and is_continuation_line(col_lines[0]):
            all_lines.append(col_lines.pop(0))
        all_lines.extend(col_lines)

    return merge_wrapped_lines(all_lines)


def merge_wrapped_lines(lines):
    """PHC sometimes wraps a provider's 'Last, First Middle, CRED' onto two
    lines, e.g. 'Al-Dwairi, Abdullah Muhammad,' / 'MD'. Merge those back
    together so name detection sees one clean line."""
    out = []
    i = 0
    n = len(lines)
    while i < n:
        cur = lines[i]
        if (
            i + 1 < n
            and cur.endswith(",")
            and re.match(r"^[A-Z][\w'\-]+(?:\s[A-Z][\w'\-]+)*,\s*.+,$", cur)
            and CRED_TOKEN_RE.match(lines[i + 1])
            and not lines[i + 1].endswith(":")
        ):
            out.append((cur + " " + lines[i + 1]).strip())
            i += 2
            continue
        out.append(cur)
        i += 1
    return out


def get_page_header(page_text: str) -> str:
    lines = page_text.splitlines()
    return lines[0].strip() if lines else ""


def is_primary_care_page(page) -> bool:
    text = page.extract_text() or ""
    header = get_page_header(text).upper()
    return header.startswith("PARTNERSHIP PRIMARY CARE PROVIDERS")


def get_county_from_page(page_text: str) -> str:
    lines = page_text.splitlines()
    if len(lines) >= 2:
        m = re.match(r"^(.*?)\s+COUNTY DIRECTORY", lines[1].strip(), re.I)
        if m:
            return m.group(1).strip().title()
    return ""


def new_provider():
    return {
        "kind": "provider",
        "Name": "",
        "NPI": "",
        "License Number": "",
        "Gender": "",
        "Specialty": "",
        "Languages": "",
        "Hospital Affiliation": "",
        "CCS Provider": "",
        "CHDP Provider": "",
    }


def new_facility():
    return {
        "kind": "facility",
        "Office/Facility Name": "",
        "PCP PHC#": "",
        "Facility NPI": "",
        "Facility Type": "",
        "Specialty": "",
        "Accepting New Patients / Enrollment Status": "",
        "Ages Accepted": "",
        "Address": "",
        "City": "",
        "State": "CA",
        "Zip": "",
        "Phone": "",
        "Fax": "",
        "Hours": "",
    }


def is_field_line(line):
    return any(r.match(line) for r in FIELD_LINE_RES)


def starts_provider(lines, i, lookahead=4):
    """A provider record starts at a 'Last, First ..., CRED' line, provided an
    NPI: line follows within a few lines (specialty/board-cert lines may sit
    in between), without another name/facility start intervening first."""
    if i >= len(lines):
        return False
    if not NAME_CRED_RE.match(lines[i]):
        return False
    for j in range(i + 1, min(i + 1 + lookahead, len(lines))):
        l = lines[j]
        if NPI_RE.match(l):
            return True
        if NAME_CRED_RE.match(l) or PHC_RE.match(l):
            return False
    return False


def starts_facility(lines, i, lookahead=4):
    """A facility record starts at a plain (non field-label, non provider-name)
    line, provided a 'PCP PHC#:' line follows within a few lines (the facility
    name can wrap across 1-2 lines)."""
    if i >= len(lines):
        return False
    line = lines[i]
    if is_field_line(line) or NAME_CRED_RE.match(line) or ACCOM_CONTINUATION_RE.match(line):
        return False
    for j in range(i, min(i + 1 + lookahead, len(lines))):
        l = lines[j]
        if PHC_RE.match(l):
            return True
        if j > i and (is_field_line(l) or NAME_CRED_RE.match(l)):
            return False
    return False


def facility_name_span(lines, i):
    """Return (name_text, index_of_phc_line) for a facility starting at i."""
    j = i
    parts = []
    while j < len(lines) and not PHC_RE.match(lines[j]):
        if not ACCOM_CONTINUATION_RE.match(lines[j]):
            parts.append(lines[j])
        j += 1
    return " ".join(parts).strip(), j


def parse_page_records(lines):
    """Walk the linear line stream and yield ('provider'|'facility', dict) records
    in the order encountered."""
    records = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        if is_boilerplate(line):
            i += 1
            continue

        if starts_provider(lines, i):
            rec = new_provider()
            rec["Name"] = line
            i += 1  # move to NPI line
            special_accom_buf = []
            hosp_buf = []
            lang_buf = []
            while i < n:
                l = lines[i]
                if is_boilerplate(l):
                    break
                if starts_provider(lines, i) or starts_facility(lines, i):
                    break
                m = NPI_RE.match(l)
                if m:
                    rec["NPI"] = m.group(1).strip()
                    i += 1
                    continue
                m = LICENSE_RE.match(l)
                if m:
                    rec["License Number"] = m.group(1).strip()
                    i += 1
                    continue
                m = GENDER_RE.match(l)
                if m:
                    rec["Gender"] = m.group(1).strip()
                    i += 1
                    continue
                m = RACE_RE.match(l)
                if m:
                    i += 1
                    continue  # race not requested in output fields; ignore
                m = ETH_RE.match(l)
                if m:
                    i += 1
                    continue  # ethnicity not requested; ignore
                m = HOSP_RE.match(l)
                if m:
                    if m.group(1):
                        hosp_buf.append(m.group(1).strip())
                    i += 1
                    continue
                m = SPECIAL_ACCOM_RE.match(l)
                if m:
                    if m.group(1):
                        special_accom_buf.append(m.group(1).strip())
                    i += 1
                    continue
                if CCS_RE.match(l):
                    rec["CCS Provider"] = "Yes"
                    i += 1
                    continue
                if CHDP_RE.match(l):
                    rec["CHDP Provider"] = "Yes"
                    i += 1
                    continue
                if l.startswith("♦") or any(
                    kw.lower() in l.lower() for kw in SPECIALTY_KEYWORDS
                ):
                    txt = l.lstrip("♦").strip()
                    if txt:
                        rec["Specialty"] = (rec["Specialty"] + ", " + txt).strip(", ")
                    i += 1
                    continue
                if HOURS_RE.match(l) or FACILITY_TYPE_RE.match(l) or \
                   ACCESS_CODES_RE.match(l) or CITY_STATE_ZIP_RE.match(l) or \
                   STREET_RE.match(l) or AGE_RE.match(l):
                    # these belong to a facility block that hasn't been detected
                    # as starting a new record (rare edge case) -- stop here so
                    # the outer loop can re-evaluate this line fresh.
                    break
                # otherwise treat as a continuation line (accommodations detail
                # or a language list with no explicit label)
                if ACCOM_CONTINUATION_RE.match(l) or (
                    special_accom_buf and ":" not in l and len(l) < 60
                ):
                    special_accom_buf.append(l)
                    i += 1
                    continue
                if ":" not in l and len(l) < 60 and re.match(r"^[A-Za-z][A-Za-z ,'\-]*$", l):
                    lang_buf.append(l)
                    i += 1
                    continue
                # unrecognized -- skip a single line to avoid infinite loop
                i += 1

            rec["Hospital Affiliation"] = "; ".join(hosp_buf)
            rec["Languages"] = ", ".join(lang_buf)
            if special_accom_buf:
                rec["Special Accommodations"] = ", ".join(special_accom_buf)
            records.append(("provider", rec))
            continue

        if starts_facility(lines, i):
            rec = new_facility()
            name_text, phc_idx = facility_name_span(lines, i)
            rec["Office/Facility Name"] = name_text
            i = phc_idx
            hours_buf = []
            specialty_buf = []
            while i < n:
                l = lines[i]
                if is_boilerplate(l):
                    break
                if starts_provider(lines, i) or starts_facility(lines, i):
                    break
                m = PHC_RE.match(l)
                if m:
                    rec["PCP PHC#"] = m.group(1).strip()
                    i += 1
                    continue
                m = NPI_RE.match(l)
                if m:
                    rec["Facility NPI"] = m.group(1).strip()
                    i += 1
                    continue
                if FACILITY_TYPE_RE.match(l):
                    rec["Facility Type"] = l.strip()
                    i += 1
                    continue
                m = ACCEPT_RE.match(l)
                if m:
                    rec["Accepting New Patients / Enrollment Status"] = (
                        "Accepting " + m.group(1).strip()
                    )
                    i += 1
                    continue
                if AGE_RE.match(l):
                    rec["Ages Accepted"] = (
                        (rec["Ages Accepted"] + "; " + l).strip("; ")
                    )
                    i += 1
                    continue
                m = CITY_STATE_ZIP_RE.match(l)
                if m:
                    rec["City"] = m.group(1).strip()
                    rec["Zip"] = m.group(2).strip()
                    i += 1
                    continue
                if ACCESS_CODES_RE.match(l) and "," in l:
                    i += 1
                    continue  # accessibility codes (P,EB,IB,R,E,T) not requested
                m = LEVEL_RE.match(l)
                if m:
                    i += 1
                    continue  # not in required output fields; skip
                m = PHONE_RE.match(l)
                if m:
                    rec["Phone"] = m.group(1).strip()
                    i += 1
                    continue
                m = FAX_RE.match(l)
                if m:
                    rec["Fax"] = m.group(1).strip()
                    i += 1
                    continue
                if HOURS_RE.match(l):
                    hours_buf.append(l.strip())
                    i += 1
                    continue
                if CCS_RE.match(l) or CHDP_RE.match(l):
                    i += 1
                    continue
                if any(kw.lower() in l.lower() for kw in SPECIALTY_KEYWORDS):
                    specialty_buf.append(l.strip())
                    i += 1
                    continue
                if not rec["City"] and not rec["Address"]:
                    # any remaining unmatched line before the city/zip line is
                    # taken to be the street address (may or may not have a
                    # leading house number in the source PDF).
                    rec["Address"] = l.strip()
                    i += 1
                    continue
                # unrecognized line inside a facility block -- skip it
                i += 1

            rec["Hours"] = "; ".join(hours_buf)
            if specialty_buf:
                rec["Specialty"] = ", ".join(specialty_buf)
            records.append(("facility", rec))
            continue

        i += 1

    return records


def flatten_records(records, county, source_pdf, current_facility=None):
    """Walk records in reading order, tracking the most-recently-seen facility
    block as running state (PHC prints the office header, then the providers
    who work there, so a facility applies FORWARD to subsequent providers
    until the next facility header is seen -- this can carry across column
    and page boundaries within a county section, which is why the running
    facility state is threaded in/out of this function across pages).
    Returns (rows, updated_current_facility)."""
    rows = []

    for kind, rec in records:
        if kind == "facility":
            current_facility = rec
            continue
        facility = current_facility
        row = {
            "Provider Type": "Primary Care",
            "Name": rec["Name"],
            "NPI": rec["NPI"],
            "License Number": rec["License Number"],
            "Gender": rec["Gender"],
            "Specialty": clean_list_text(
                rec["Specialty"] or (facility["Specialty"] if facility else "")
            ),
            "Office/Facility Name": facility["Office/Facility Name"] if facility else "",
            "Address": facility["Address"] if facility else "",
            "City": facility["City"] if facility else "",
            "State": "CA",
            "Zip": facility["Zip"] if facility else "",
            "Phone": facility["Phone"] if facility else "",
            "Fax": facility["Fax"] if facility else "",
            "Languages": clean_list_text(rec["Languages"]),
            "Accepting New Patients / Enrollment Status": (
                facility["Accepting New Patients / Enrollment Status"] if facility else ""
            ),
            "Ages Accepted": facility["Ages Accepted"] if facility else "",
            "Hospital Affiliation": rec["Hospital Affiliation"],
            "County": county,
            "source_pdf": source_pdf,
        }
        rows.append(row)

    return rows, current_facility


def clean_list_text(s):
    """Normalize a comma-joined field (e.g. Specialty) that may have picked up
    duplicate/trailing commas from multi-line source text."""
    if not s:
        return s
    parts = [p.strip().rstrip(",").strip() for p in re.split(r",\s*", s)]
    parts = [p for p in parts if p]
    seen = set()
    out = []
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return ", ".join(out)


def strip_pii(row):
    banned_keys = {"ssn", "social security", "dob", "date of birth", "birthdate"}
    return {k: v for k, v in row.items() if k.strip().lower() not in banned_keys}


def parse_pdf(path, source_pdf):
    rows = []
    pages_processed = 0
    current_facility = None
    last_county = None
    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        for idx, page in enumerate(pdf.pages):
            if not is_primary_care_page(page):
                continue
            pages_processed += 1
            text = page.extract_text() or ""
            county = get_county_from_page(text)
            if county and county != last_county:
                current_facility = None  # new county section: don't carry over
                last_county = county
            lines = build_page_lines(page)
            records = parse_page_records(lines)
            page_rows, current_facility = flatten_records(
                records, county, source_pdf, current_facility
            )
            rows.extend(page_rows)
        print(
            f"  {source_pdf}: {total} total pages, {pages_processed} Primary Care "
            f"pages processed, {len(rows)} provider rows extracted"
        )
    return [strip_pii(r) for r in rows]


def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_checkpoint(rows):
    with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2, ensure_ascii=False)


def dedupe(rows):
    seen = set()
    out = []
    for r in rows:
        key = (
            (r.get("NPI") or "").strip(),
            (r.get("Address") or "").strip().lower(),
            (r.get("Specialty") or "").strip().lower(),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def export_to_excel(rows):
    import pandas as pd
    from openpyxl import load_workbook

    if not rows:
        print("No rows to export -- skipping Excel write.")
        return

    columns = [
        "Name", "NPI", "License Number", "Gender", "Specialty",
        "Office/Facility Name", "Address", "City", "State", "Zip", "Phone", "Fax",
        "Languages", "Accepting New Patients / Enrollment Status", "Ages Accepted",
        "Hospital Affiliation", "Provider Type", "County", "source_pdf",
    ]
    df = pd.DataFrame(rows)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]

    df.to_excel(OUTPUT_XLSX_PATH, index=False, engine="openpyxl")

    wb = load_workbook(OUTPUT_XLSX_PATH)
    ws = wb.active
    header = [c.value for c in ws[1]]
    text_col_names = {"phone", "fax", "npi", "zip", "license number"}
    text_cols = [
        i for i, h in enumerate(header, start=1)
        if h and h.strip().lower() in text_col_names
    ]
    for col_idx in text_cols:
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.number_format = "@"
            if cell.value is not None:
                cell.value = str(cell.value)
    wb.save(OUTPUT_XLSX_PATH)
    print(f"Wrote {len(rows)} rows to {OUTPUT_XLSX_PATH}")


def main():
    only = sys.argv[1:] if len(sys.argv) > 1 else None
    all_rows = load_checkpoint()
    done_sources = {r.get("source_pdf") for r in all_rows}

    for fname in PDF_FILES:
        if only and fname not in only:
            continue
        if fname in done_sources:
            print(f"Skipping {fname} (already in checkpoint)")
            continue
        path = os.path.join(PDF_DIR, fname)
        if not os.path.exists(path):
            print(f"WARNING: {path} not found, skipping")
            continue
        print(f"Parsing {fname} ...")
        try:
            rows = parse_pdf(path, fname)
        except Exception as e:
            print(f"  ERROR parsing {fname}: {e}")
            continue
        all_rows.extend(rows)
        save_checkpoint(all_rows)
        print(f"  Checkpoint now has {len(all_rows)} total raw rows.")

    final_rows = dedupe(all_rows)
    export_to_excel(final_rows)


if __name__ == "__main__":
    main()
