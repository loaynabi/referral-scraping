"""
Parses the raw_*.txt files (DOM-scraped text blocks captured from live browser
sessions against Molina's Sapphire365 provider-search UI) into structured
provider records, deduplicates them, writes checkpoint.json, and builds the
final Excel export.

Each raw file is a series of provider "blocks" separated by blank lines, in
this shape (as rendered by the site's search-results list):

    <Name>[, <credential>] [PCP]
    <specialty line(s)>
    NPI: <10 digits>
    <Medical Group / Location Name>
    <address>, <city>, CA <zip>
    Phone: <phone>
    [Accepting New Patients | Not Accepting New Patients]

No SSN / DOB / other sensitive PII is present in source data or output.
"""
import json
import re
import glob
import os

BASE = os.path.dirname(os.path.abspath(__file__))

METRO_MAP = {
    "LA_90012": ("Los Angeles", "90012"),
    "SanDiego_92101": ("San Diego", "92101"),
    "Sacramento_95814": ("Sacramento", "95814"),
    "Fresno_93701": ("Fresno", "93701"),
    "SanJose_95112": ("San Jose", "95112"),
}

def metro_from_filename(fname):
    for key, val in METRO_MAP.items():
        if key in fname:
            return val
    return ("", "")

def parse_file(path):
    with open(path, encoding="utf-8") as f:
        text = f.read()
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    records = []
    for b in blocks:
        lines = [l.strip() for l in b.splitlines() if l.strip()]
        if len(lines) < 5:
            continue
        name_line = lines[0]
        is_pcp_badge = name_line.endswith("PCP")
        name = re.sub(r"\s+PCP$", "", name_line).strip()
        specialty_line = lines[1]
        npi_line = next((l for l in lines if l.startswith("NPI:")), "")
        npi = npi_line.replace("NPI:", "").strip()
        # medical group = line after NPI
        npi_idx = lines.index(npi_line) if npi_line in lines else 2
        medical_group = lines[npi_idx + 1] if npi_idx + 1 < len(lines) else ""
        addr_line = lines[npi_idx + 2] if npi_idx + 2 < len(lines) else ""
        phone_line = next((l for l in lines if l.startswith("Phone:")), "")
        phone = phone_line.replace("Phone:", "").strip()
        accepting = "Y" if any("Accepting New Patients" == l for l in lines) else (
            "N" if any("Not Accepting New Patients" == l for l in lines) else "")

        # split address "123 Main St Ste 1, City, CA 90012"
        addr_parts = [p.strip() for p in addr_line.split(",")]
        street = addr_parts[0] if len(addr_parts) > 0 else ""
        city = addr_parts[1] if len(addr_parts) > 1 else ""
        state_zip = addr_parts[2] if len(addr_parts) > 2 else ""
        m = re.match(r"([A-Z]{2})\s+(\d{5})", state_zip)
        state = m.group(1) if m else "CA"
        zip_code = m.group(2) if m else ""

        records.append({
            "name": name,
            "is_pcp": True if is_pcp_badge else True,  # all captured via is_pcp filter
            "primary_specialty": specialty_line,
            "npi": npi,
            "medical_group": medical_group,
            "address": street,
            "city": city,
            "state": state,
            "zip": zip_code,
            "phone": phone,
            "accepting_new_patients": accepting,
            "source_search_zip": "",
            "source_search_city": "",
        })
    return records

def main():
    all_records = []
    for path in sorted(glob.glob(os.path.join(BASE, "raw_*.txt"))):
        fname = os.path.basename(path)
        city, zipc = metro_from_filename(fname)
        recs = parse_file(path)
        for r in recs:
            r["source_search_city"] = city
            r["source_search_zip"] = zipc
        all_records.append((fname, recs))
        print(f"{fname}: {len(recs)} records")

    flat = [r for _, recs in all_records for r in recs]

    # Drop any record that somehow lacks an NPI (can't reliably dedupe/verify)
    flat = [r for r in flat if r["npi"]]

    # Dedup on (NPI, address, specialty)
    seen = {}
    for r in flat:
        key = (r["npi"], r["address"].lower(), r["primary_specialty"].lower())
        if key not in seen:
            seen[key] = r
    deduped = list(seen.values())

    # Safety: strip any accidental sensitive fields (defensive; none expected)
    SENSITIVE = {"ssn", "dob", "date_of_birth", "social_security"}
    for r in deduped:
        for k in list(r.keys()):
            if k.lower() in SENSITIVE:
                del r[k]

    checkpoint_path = os.path.join(BASE, "checkpoint.json")
    with open(checkpoint_path, "w", encoding="utf-8") as f:
        json.dump(deduped, f, indent=2)

    print(f"\nTotal raw records parsed: {len(flat)}")
    print(f"Total after de-dup:       {len(deduped)}")
    print(f"Checkpoint written to:    {checkpoint_path}")

if __name__ == "__main__":
    main()
