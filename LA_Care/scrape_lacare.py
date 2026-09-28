"""
L.A. Care Health Plan (Medi-Cal) — Primary Care Provider directory scraper.

Data source: L.A. Care's public Provider Directory JSON API (no auth/session
required, confirmed stateless):
    POST https://providers.lacare.org/api/ProviderDirectoryV1/JSON/Practitioners/Search
        ?serviceid=c2aba49a-f2b8-4191-96ad-2c866d4fec60
        &siteid=186a3f3f-a8c2-4334-95ad-3f239ff14b74

The API is radius/geo based. A single query centered on the LA County
centroid with radius=100 (miles) returns the full PCP population for the
whole county (verified: TotalRecordCount is stable at ~198 for radius >= 30,
i.e. the county is already fully covered well inside 100 miles). Filter
`providertype:PCP` uses the API's own PCP/Specialist distinction
(ProviderTypes[].Code == "PCP" vs "SPEC" etc.) — this is the equivalent of
the contractTypePcpSpc field used on prior sites.

Checkpointing: progress is saved to checkpoint.json after every page so a
crash/interrupt loses at most one page (count=100 per page).
"""
import json
import random
import time
from pathlib import Path

import requests

BASE_DIR = Path(__file__).parent
CHECKPOINT_FILE = BASE_DIR / "checkpoint.json"
OUTPUT_FILE = BASE_DIR / "LA_Care_Providers.xlsx"

API_URL = "https://providers.lacare.org/api/ProviderDirectoryV1/JSON/Practitioners/Search"
PARAMS = {
    "serviceid": "c2aba49a-f2b8-4191-96ad-2c866d4fec60",
    "siteid": "186a3f3f-a8c2-4334-95ad-3f239ff14b74",
}
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json",
}

# LA County centroid; radius=100mi comfortably covers the whole service area
# (verified TotalRecordCount stops growing well before this radius).
CENTER_LAT = 34.0522
CENTER_LON = -118.2437
RADIUS_MILES = 100
PAGE_SIZE = 100  # server errors (500) above ~100-149


def fetch_page(start: int) -> dict:
    body = {
        "start": start,
        "count": PAGE_SIZE,
        "searchModifiers": (
            f"providertype:PCP|lat:{CENTER_LAT}|lon:{CENTER_LON}|"
            f"radius:{RADIUS_MILES}|location:Los Angeles County, CA|"
            "sortfield:Distance|sortdir:ascending"
        ),
        "flatten": True,
        "language": "en",
    }
    for attempt in range(5):
        try:
            r = requests.post(API_URL, params=PARAMS, json=body, headers=HEADERS, timeout=30)
        except requests.RequestException as e:
            print(f"  request error ({e}); retrying in {2 ** attempt}s")
            time.sleep(2 ** attempt)
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code in (401, 403, 429):
            print(f"  got {r.status_code} — pausing, possible block/rate-limit")
            time.sleep(10 * (attempt + 1))
            continue
        print(f"  unexpected status {r.status_code}; retrying in {2 ** attempt}s")
        time.sleep(2 ** attempt)
    raise RuntimeError(f"failed to fetch start={start} after retries")


def load_checkpoint() -> dict:
    if CHECKPOINT_FILE.exists():
        return json.loads(CHECKPOINT_FILE.read_text())
    return {"next_start": 1, "records": [], "total": None, "done": False}


def save_checkpoint(state: dict):
    CHECKPOINT_FILE.write_text(json.dumps(state, indent=2))


SENSITIVE_KEYS = {"ssn", "dob", "dateofbirth", "birthdate", "socialsecuritynumber"}


def strip_sensitive(d: dict) -> dict:
    return {k: v for k, v in d.items() if k.lower() not in SENSITIVE_KEYS}


def flatten_records(raw_records: list) -> list:
    rows = []
    for rec in raw_records:
        rec = strip_sensitive(rec)
        name = rec.get("FullName") or f"{rec.get('FirstName','')} {rec.get('LastName','')}".strip()
        npi = rec.get("NPIN", "")
        specialties = rec.get("SpecialtiesCombined", "")
        provider_types = rec.get("ProviderTypesCombined", "")
        languages = rec.get("LanguagesCombined", "")
        networks = rec.get("NetworksCombined", "")
        facilities = rec.get("Facilities") or [{}]
        for fac in facilities:
            fac = strip_sensitive(fac)
            rows.append({
                "Name": name,
                "Phone": fac.get("Phone1") or fac.get("PhoneCombined", ""),
                "Fax": "",  # not exposed by this API
                "Specialty": specialties,
                "Provider Type (PCP/Specialist)": provider_types,
                "Office Name": fac.get("FacilityName", ""),
                "Address": fac.get("Address1", ""),
                "City": fac.get("City", ""),
                "State": fac.get("State", ""),
                "Zip": fac.get("Zip", ""),
                "NPI": npi,
                "Plan / Network": networks,
                "Accepting New Patients": fac.get("AcceptingNewPatients", ""),
                "Languages Spoken": languages,
                "Gender": rec.get("Gender", ""),
                "Distance (mi from LA County center)": rec.get("Distance", ""),
            })
    return rows


def dedupe(rows: list) -> list:
    seen = set()
    out = []
    for r in rows:
        key = (r["NPI"], r["Address"], r["Specialty"])
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def main():
    state = load_checkpoint()
    if state["done"]:
        print("Checkpoint marked done; delete checkpoint.json to re-run.")
    else:
        while True:
            start = state["next_start"]
            print(f"Fetching start={start} ...")
            data = fetch_page(start)
            total = data.get("TotalRecordCount")
            records = data.get("Records") or []
            if total is not None:
                state["total"] = total
            state["records"].extend(records)
            print(f"  got {len(records)} records (running total {len(state['records'])} / {state['total']})")
            save_checkpoint(state)
            if not records or (state["total"] and start + PAGE_SIZE > state["total"]):
                state["done"] = True
                save_checkpoint(state)
                break
            state["next_start"] = start + PAGE_SIZE
            time.sleep(random.uniform(1, 3))

    rows = flatten_records(state["records"])
    rows = dedupe(rows)
    print(f"Total unique PCP rows: {len(rows)}")

    import pandas as pd

    df = pd.DataFrame(rows)
    text_cols = ["Phone", "Fax", "NPI", "Zip"]
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].astype(str)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="PCP Providers")
        ws = writer.sheets["PCP Providers"]
        for col in text_cols:
            if col in df.columns:
                col_idx = df.columns.get_loc(col) + 1
                for cell in ws.iter_rows(min_col=col_idx, max_col=col_idx, min_row=2):
                    cell[0].number_format = "@"

    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
