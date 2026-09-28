"""
CalOptima Health — Primary Care Provider directory scraper.

Data source: CalOptima's Sitecore Search (Sitecore Discover) backend, a
stateless public JSON API confirmed to require no auth/cookies:
    POST https://discover.sitecorecloud.io/discover/v2/87111946

Filtered server-side to provider_directory_type == "Primary Care Provider"
(CalOptima's own PCP/Specialist distinction). No plan/network filter is
applied so the broadest complete PCP list is pulled (covers Medi-Cal,
OneCare, etc. — the "Plan / Network" column shows what's actually offered
per provider/location).

Checkpointing: progress saved to checkpoint.json after every page (limit=100
providers per page) so a crash loses at most one page.
"""
import json
import random
import time
from pathlib import Path

import requests

BASE_DIR = Path(__file__).parent
CHECKPOINT_FILE = BASE_DIR / "checkpoint.json"
OUTPUT_FILE = BASE_DIR / "CalOptima_Providers.xlsx"

API_URL = "https://discover.sitecorecloud.io/discover/v2/87111946"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36", "Content-Type": "application/json"}
PAGE_SIZE = 100


def fetch_page(offset: int) -> dict:
    body = {
        "context": {"locale": {"language": "en", "country": "us"}, "page": {"uri": "/en/providers"}},
        "widget": {"items": [{
            "rfk_id": "rfkid_7",
            "search": {
                "content": {},
                "offset": offset,
                "limit": PAGE_SIZE,
                "sort": {"value": [{"name": "relevance"}]},
                "filter": {"type": "and", "filters": [
                    {"type": "anyOf", "name": "provider_directory_type", "values": ["Primary Care Provider"]},
                ]},
            },
            "entity": "content",
            "sources": ["1100434"],
        }]},
    }
    for attempt in range(5):
        try:
            r = requests.post(API_URL, json=body, headers=HEADERS, timeout=30)
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
    raise RuntimeError(f"failed to fetch offset={offset} after retries")


def load_checkpoint() -> dict:
    if CHECKPOINT_FILE.exists():
        return json.loads(CHECKPOINT_FILE.read_text())
    return {"next_offset": 0, "records": [], "total": None, "done": False}


def save_checkpoint(state: dict):
    CHECKPOINT_FILE.write_text(json.dumps(state, indent=2))


SENSITIVE_SUBSTRINGS = ("ssn", "social_security", "dob", "date_of_birth", "birthdate", "birth_date")


def strip_sensitive(d: dict) -> dict:
    return {k: v for k, v in d.items() if not any(s in k.lower() for s in SENSITIVE_SUBSTRINGS)}


def joinlist(v):
    if isinstance(v, list):
        return ", ".join(str(x) for x in v if x not in (None, ""))
    return v or ""


def flatten_records(raw_records: list) -> list:
    rows = []
    for rec in raw_records:
        rec = strip_sensitive(rec)
        name = rec.get("provider_full_name") or rec.get("name", "")
        npi = rec.get("provider_npi", "")
        specialty = joinlist(rec.get("provider_specialty") or rec.get("provider_board_certified"))
        license_num = rec.get("provider_license", "")
        directory_type = joinlist(rec.get("provider_directory_type"))
        languages = joinlist(rec.get("provider_language_spoken"))
        accepting = joinlist(rec.get("provider_accepting_new_patients"))
        locations = rec.get("provider_locations") or [{}]
        for loc in locations:
            loc = strip_sensitive(loc)
            health_networks = loc.get("healthNetworks") or []
            plan_network = ", ".join(
                f"{hn.get('name','')} ({hn.get('lineOfBusiness','')})" for hn in health_networks if hn.get("name")
            )
            rows.append({
                "Name": name,
                "Phone": loc.get("daytimePhone", ""),
                "Fax": loc.get("fax", ""),
                "Specialty": joinlist(loc.get("specialty")) or specialty,
                "Provider Type (PCP/Specialist)": joinlist(loc.get("directoryTypeHeader")) or directory_type,
                "Office Name": loc.get("medicalGroupName", ""),
                "Address": ", ".join(x for x in [loc.get("addressLine1", ""), loc.get("addressLine2", "")] if x),
                "City": loc.get("city", ""),
                "State": loc.get("state", ""),
                "Zip": loc.get("zip", ""),
                "NPI": npi,
                "Plan / Network": plan_network,
                "Accepting New Patients": joinlist(loc.get("acceptingNewPatients")) or accepting,
                "License Number": license_num,
                "Languages Spoken": languages,
                "After Hours Phone": loc.get("afterHoursPhone", ""),
                "Facility Affiliations": joinlist(rec.get("provider_facility_affiliations")),
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
            offset = state["next_offset"]
            print(f"Fetching offset={offset} ...")
            data = fetch_page(offset)
            widget = (data.get("widgets") or [{}])[0]
            total = widget.get("total_item")
            records = widget.get("content") or []
            if total is not None:
                state["total"] = total
            state["records"].extend(records)
            print(f"  got {len(records)} records (running total {len(state['records'])} / {state['total']})")
            save_checkpoint(state)
            if not records or (state["total"] and offset + PAGE_SIZE >= state["total"]):
                state["done"] = True
                save_checkpoint(state)
                break
            state["next_offset"] = offset + PAGE_SIZE
            time.sleep(random.uniform(1, 3))

    rows = flatten_records(state["records"])
    rows = dedupe(rows)
    print(f"Total unique PCP rows: {len(rows)}")

    import pandas as pd

    df = pd.DataFrame(rows)
    text_cols = ["Phone", "Fax", "NPI", "Zip", "After Hours Phone", "License Number"]

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
