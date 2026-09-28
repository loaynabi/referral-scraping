#!/usr/bin/env python3
"""
Sharp Health Plan / HealthTrio California Provider Directory Scraper
======================================================================

Scrapes the public consumer provider-directory search results at
sharp.healthtrioconnect.com and exports them to Excel + CSV.

Design notes (see chat for full selector analysis):
  - Uses a VISIBLE, persistent Chrome profile via Selenium. Cookies/session
    state live inside that Chrome profile directory on disk (managed by
    Chrome itself) -- this script never reads, prints, stores, or otherwise
    touches cookie values.
  - Never attempts to detect-and-solve Cloudflare/CAPTCHA challenges. When a
    verification screen is detected, the script pauses and waits for you to
    solve it by hand in the visible browser window, then continues in the
    SAME browser session.
  - Always follows the real "Next Page" link found in the loaded page,
    rather than constructing startRow URLs itself.
  - Saves progress (Excel + CSV + JSON state) after every results page, so
    Ctrl+C or a crash never loses more than one page of work.
  - Detail-page fields (fax, license number, languages, office hours,
    additional addresses) are extracted with GENERIC, best-effort text
    pattern matching, because no sample detail-page HTML was available to
    build confirmed selectors against. The first couple of detail pages are
    dumped to debug/ for you to send back so the extraction can be
    tightened. Every other required field is covered directly by the
    results-page cards and is not "best effort".

Usage (see README block at bottom / chat message for exact PowerShell
commands):
    python sharp_scraper.py
    python sharp_scraper.py --skip-details
    python sharp_scraper.py --fresh
"""

import argparse
import csv
import json
import math
import random
import re
import sys
import time
import traceback
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

ORIGINAL_TARGET_URL = (
    "https://sharp.healthtrioconnect.com/public-app/consumer/provdir/"
    "search.page?isEntry=entry&startRow=0&sortType=NAMEASC&provType=P&addressState=ca"
)

OUTPUT_DIR = Path("./output")
DEBUG_DIR = OUTPUT_DIR / "debug"
DEFAULT_PROFILE_DIR = Path("./chrome_profile").resolve()

XLSX_PATH = OUTPUT_DIR / "sharp_healthtrio_california_providers.xlsx"
CSV_PATH = OUTPUT_DIR / "sharp_healthtrio_california_providers.csv"
LOG_PATH = OUTPUT_DIR / "scrape_log.txt"
STATE_PATH = OUTPUT_DIR / "scrape_progress.json"

PAGE_DELAY_RANGE = (10, 20)     # seconds, between results pages
DETAIL_DELAY_RANGE = (3, 5)     # seconds, between provider detail pages
RESULTS_WAIT_TIMEOUT = 45       # seconds to wait for a results page to render
DETAIL_WAIT_TIMEOUT = 30
MAX_TRANSIENT_RETRIES = 4
STARTUP_RETRY_DELAY = 5         # seconds, doubles each retry (exponential backoff)

HEADERS = [
    "Provider Name", "Specialty", "NPI", "Phone", "Fax",
    "Medical Group", "Practice/Facility",
    "Street Address", "City", "State", "ZIP Code",
    "Accepting New Patients", "Accepting New Medicare Patients",
    "License Number", "Languages", "Office Hours",
    "Additional Specialties", "Additional Addresses",
    "Provider Detail URL", "Source URL", "Scraped Timestamp",
    "Detail Page Reviewed",
]

VERIFICATION_INDICATORS = [
    "just a moment", "checking your browser", "attention required",
    "cf-browser-verification", "cf-chl", "verify you are human",
    "enable javascript and cookies to continue", "ray id",
    "please stand by, while we are checking your browser",
]

ERROR_INDICATORS = [
    "internal server error", "503 service unavailable",
    "service temporarily unavailable", "an unexpected error has occurred",
    "the page you requested is not available", "bad gateway",
    "504 gateway", "502 bad gateway",
]


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def log(msg: str, also_print: bool = True):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    if also_print:
        print(line)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def log_error(context: str, url: str, start_row, error: str):
    log(f"ERROR | context={context} | url={url} | startRow={start_row} | {error}")


# ---------------------------------------------------------------------------
# Data row
# ---------------------------------------------------------------------------

@dataclass
class ProviderRow:
    provider_name: str = ""
    specialty: str = ""
    npi: str = ""
    phone: str = ""
    fax: str = ""
    medical_group: str = ""
    practice_facility: str = ""
    street_address: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    accepting_new_patients: str = ""
    accepting_new_medicare: str = ""
    license_number: str = ""
    languages: str = ""
    office_hours: str = ""
    additional_specialties: str = ""
    additional_addresses: str = ""
    provider_detail_url: str = ""
    source_url: str = ""
    scraped_timestamp: str = ""
    detail_reviewed: str = "N"

    def to_row(self):
        return [
            self.provider_name, self.specialty, self.npi, self.phone, self.fax,
            self.medical_group, self.practice_facility,
            self.street_address, self.city, self.state, self.zip_code,
            self.accepting_new_patients, self.accepting_new_medicare,
            self.license_number, self.languages, self.office_hours,
            self.additional_specialties, self.additional_addresses,
            self.provider_detail_url, self.source_url, self.scraped_timestamp,
            self.detail_reviewed,
        ]

    def dedup_key(self):
        if self.npi:
            return f"npi:{self.npi}"
        return f"url:{self.provider_detail_url}|{self.provider_name}|{self.street_address}"

    def needs_detail_enrichment(self):
        return self.detail_reviewed != "Y"


# ---------------------------------------------------------------------------
# Persistence: state, workbook, csv
# ---------------------------------------------------------------------------

def load_state():
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            log(f"WARNING: could not read state file ({e}); starting fresh state.")
    return {"last_start_row": -10, "total_count": None}


def save_state(state: dict):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def load_existing_providers():
    """Resume support: read any existing workbook back into memory."""
    providers = {}
    if not XLSX_PATH.exists():
        return providers
    try:
        wb = load_workbook(XLSX_PATH)
        ws = wb["Providers"]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        for r in rows:
            if not r or all(v is None for v in r):
                continue
            vals = [("" if v is None else str(v)) for v in r]
            while len(vals) < len(HEADERS):
                vals.append("")
            row = ProviderRow(*vals[: len(HEADERS)])
            providers[row.dedup_key()] = row
        log(f"Resumed: loaded {len(providers)} existing providers from {XLSX_PATH}")
    except Exception as e:
        log(f"WARNING: could not load existing workbook ({e}); starting with empty set.")
    return providers


def save_all(providers: dict):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    _save_xlsx(providers)
    _save_csv(providers)


def _save_xlsx(providers: dict):
    wb = Workbook()
    ws = wb.active
    ws.title = "Providers"
    ws.append(HEADERS)
    for row in sorted(providers.values(), key=lambda r: (r.provider_name or "")):
        ws.append(row.to_row())

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = [26, 20, 14, 16, 16, 30, 30, 26, 16, 8, 12, 16, 18, 16, 18, 16, 22, 24, 46, 46, 20, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    tmp = XLSX_PATH.with_suffix(".xlsx.tmp")
    wb.save(tmp)
    tmp.replace(XLSX_PATH)


def _save_csv(providers: dict):
    tmp = CSV_PATH.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        for row in sorted(providers.values(), key=lambda r: (r.provider_name or "")):
            writer.writerow(row.to_row())
    tmp.replace(CSV_PATH)


# ---------------------------------------------------------------------------
# Browser setup (no anti-detection / stealth tricks -- deliberately)
# ---------------------------------------------------------------------------

def build_driver(profile_dir: Path):
    profile_dir.mkdir(parents=True, exist_ok=True)
    options = Options()
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.add_argument("--start-maximized")
    # Deliberately NOT headless, and deliberately no automation-hiding flags
    # (no excludeSwitches, no disable-blink-features=AutomationControlled,
    # no CDP fingerprint patching). Verification must be solvable by a human,
    # and this script does not try to make Chrome look non-automated.
    driver = webdriver.Chrome(options=options)
    return driver


# ---------------------------------------------------------------------------
# Page-state detection
# ---------------------------------------------------------------------------

def page_signal(driver):
    """Return one of: 'ok', 'verification', 'error', 'empty', 'unknown'."""
    try:
        title = (driver.title or "").lower()
        html = driver.page_source or ""
        html_lower = html.lower()
    except WebDriverException:
        return "unknown"

    if any(ind in title or ind in html_lower for ind in VERIFICATION_INDICATORS):
        return "verification"
    if any(ind in html_lower for ind in ERROR_INDICATORS):
        return "error"

    soup = BeautifulSoup(html, "lxml")
    if soup.select("ul.providerList li.result_set"):
        return "ok"
    if soup.select_one("div.serverPaginate") and not soup.select_one("ul.providerList"):
        return "empty"
    return "unknown"


def wait_for_page(driver, timeout=RESULTS_WAIT_TIMEOUT):
    deadline = time.time() + timeout
    while time.time() < deadline:
        sig = page_signal(driver)
        if sig != "unknown":
            return sig
        time.sleep(1)
    return "timeout"


def handle_verification(driver):
    print("\n" + "=" * 72)
    print("VERIFICATION REQUIRED")
    print("A Cloudflare / bot-check screen appeared in the browser window.")
    print("This script will NOT attempt to solve it automatically.")
    print("Please complete it manually in the visible Chrome window.")
    print("When the provider search results are showing again, come back")
    print("here and press ENTER to resume scraping.")
    print("=" * 72 + "\n")
    try:
        input()
    except EOFError:
        pass
    return wait_for_page(driver, timeout=90)


# ---------------------------------------------------------------------------
# Navigation with retries (transient errors only -- never for verification)
# ---------------------------------------------------------------------------

def safe_get(driver, url, context="navigate"):
    delay = STARTUP_RETRY_DELAY
    for attempt in range(1, MAX_TRANSIENT_RETRIES + 1):
        try:
            driver.get(url)
            return True
        except (WebDriverException, TimeoutException) as e:
            log_error(context, url, None, f"navigation error attempt {attempt}: {e}")
            if attempt == MAX_TRANSIENT_RETRIES:
                return False
            time.sleep(delay)
            delay *= 2
    return False


# ---------------------------------------------------------------------------
# Parsing: results page
# ---------------------------------------------------------------------------

NPI_RE = re.compile(r"NPI:\s*(\d+)", re.IGNORECASE)
PHONE_RE = re.compile(r"Phone:\s*(.+)", re.IGNORECASE)
CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.*?),\s*(?P<state>[A-Za-z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$")
COUNT_RE = re.compile(r"(\d+)\s*-\s*(\d+)\s+of\s+(\d+)", re.IGNORECASE)


def parse_accepting(card):
    result = {"accepting_new_patients": "", "accepting_new_medicare": ""}
    accept_span = card.select_one("span.specialty.printSection.acceptingNewPatients")
    if not accept_span:
        return result
    for container in accept_span.select("div.resultSelectionLabelContainer"):
        label = container.select_one("span[title]")
        icon = container.select_one("span.checkMark, span.xMark")
        if not label:
            continue
        title = label.get("title", "")
        is_yes = icon is not None and "checkMark" in icon.get("class", [])
        if "medicare" in title.lower():
            result["accepting_new_medicare"] = "Yes" if is_yes else "No"
        else:
            result["accepting_new_patients"] = "Yes" if is_yes else "No"
    return result


def parse_card(card, current_page_url, source_url):
    row = ProviderRow()
    row.source_url = source_url
    row.scraped_timestamp = datetime.now(timezone.utc).isoformat()

    name_a = card.select_one("div.provider_meta h3.printSection a.provider_name")
    if not name_a:
        raise ValueError("could not find provider name link in card")
    row.provider_name = name_a.get_text(strip=True)
    href = name_a.get("href", "")
    row.provider_detail_url = urljoin(current_page_url, href)

    spans = card.select("div.provider_meta > span.specialty.printSection")
    specialty_spans, npi_span = [], None
    for s in spans:
        classes = s.get("class", [])
        if "acceptingNewPatients" in classes:
            continue
        text = s.get_text(strip=True)
        if text.upper().startswith("PROVIDER NPI"):
            npi_span = s
        else:
            specialty_spans.append(text)
    if npi_span:
        m = NPI_RE.search(npi_span.get_text(strip=True))
        if m:
            row.npi = m.group(1)
    if specialty_spans:
        row.specialty = specialty_spans[0]
        if len(specialty_spans) > 1:
            row.additional_specialties = "; ".join(specialty_spans[1:])

    accepting = parse_accepting(card)
    row.accepting_new_patients = accepting["accepting_new_patients"]
    row.accepting_new_medicare = accepting["accepting_new_medicare"]

    loc = card.select_one('div.location[tabindex="0"]')
    if loc:
        direct_divs = loc.find_all("div", recursive=False)
        if len(direct_divs) >= 1:
            row.medical_group = direct_divs[0].get_text(strip=True)
        if len(direct_divs) >= 2:
            row.practice_facility = direct_divs[1].get_text(strip=True)

        addr = loc.select_one("address.printSection")
        if addr:
            top_spans = addr.find_all("span", recursive=False)
            if top_spans:
                nested = top_spans[0].find_all("span", recursive=False)
                if len(nested) >= 1:
                    row.street_address = nested[0].get_text(strip=True)
                if len(nested) >= 2:
                    csz = nested[1].get_text(strip=True)
                    m = CITY_STATE_ZIP_RE.match(csz)
                    if m:
                        row.city = m.group("city")
                        row.state = m.group("state").upper()
                        row.zip_code = m.group("zip")
                    else:
                        row.city = csz
            if len(top_spans) >= 2:
                phone_text = top_spans[1].get_text(strip=True)
                m = PHONE_RE.search(phone_text)
                row.phone = m.group(1).strip() if m else phone_text

    return row


def parse_results_page(html, current_page_url, source_url):
    soup = BeautifulSoup(html, "lxml")
    cards_html = soup.select("ul.providerList li.result_set")

    rows, errors = [], []
    for card in cards_html:
        try:
            rows.append(parse_card(card, current_page_url, source_url))
        except Exception as e:
            errors.append(str(e))
            log_error("parse_card", current_page_url, None, f"{e}")

    pag = soup.select_one("div.serverPaginate")
    total_count, first_shown, next_href = None, None, None
    if pag:
        count_el = pag.select_one("li.paginationCount")
        if count_el:
            m = COUNT_RE.search(count_el.get_text(" ", strip=True))
            if m:
                first_shown, _, total_count = int(m.group(1)), int(m.group(2)), int(m.group(3))
        next_a = pag.select_one('a[aria-label="Next Page"]')
        if next_a:
            classes = " ".join(next_a.get("class", []))
            href = next_a.get("href")
            if href and "disabled" not in classes.lower():
                next_href = urljoin(current_page_url, href)

    return {
        "rows": rows,
        "errors": errors,
        "total_count": total_count,
        "first_shown": first_shown,
        "next_href": next_href,
    }


# ---------------------------------------------------------------------------
# Detail page enrichment (best-effort / generic -- see module docstring)
# ---------------------------------------------------------------------------

DETAIL_PATTERNS = {
    "fax": re.compile(r"Fax[:\s]+([\(\)0-9\-\s\.]{7,20})", re.IGNORECASE),
    "license_number": re.compile(r"License(?:\s*(?:No\.?|Number))?[:\s]+([A-Za-z0-9\-]+)", re.IGNORECASE),
    "languages": re.compile(r"Language[s]?(?: Spoken)?[:\s]+([^\n]+)", re.IGNORECASE),
    "office_hours": re.compile(r"(?:Office\s*)?Hours[:\s]+([^\n]+)", re.IGNORECASE),
}

_debug_dump_count = 0


def extract_detail_fields(html):
    soup = BeautifulSoup(html, "lxml")
    text = soup.get_text("\n", strip=True)
    out = {"fax": "", "license_number": "", "languages": "", "office_hours": ""}
    for key, pattern in DETAIL_PATTERNS.items():
        m = pattern.search(text)
        if m:
            out[key] = m.group(1).strip()
    return out


def maybe_dump_debug_html(html, provider_id):
    global _debug_dump_count
    if _debug_dump_count >= 2:
        return
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    path = DEBUG_DIR / f"detail_debug_{_debug_dump_count + 1}_{provider_id or 'unknown'}.html"
    try:
        path.write_text(html, encoding="utf-8")
        log(f"Saved a detail-page sample for review: {path}")
    except Exception as e:
        log(f"WARNING: could not save debug detail page: {e}")
    _debug_dump_count += 1


def enrich_with_detail_pages(driver, providers: dict, state: dict, stop_flag):
    to_visit = [r for r in providers.values() if r.needs_detail_enrichment() and r.provider_detail_url]
    if not to_visit:
        log("No providers need detail-page enrichment.")
        return

    log(f"Starting detail-page enrichment for {len(to_visit)} providers "
        f"(sequential, {DETAIL_DELAY_RANGE[0]}-{DETAIL_DELAY_RANGE[1]}s delay each).")

    for i, row in enumerate(to_visit, start=1):
        if stop_flag["stop"]:
            log("Stop requested -- ending detail-page enrichment early.")
            break

        parsed = urlparse(row.provider_detail_url)
        provider_id = parse_qs(parsed.query).get("providerID", [""])[0]

        ok = safe_get(driver, row.provider_detail_url, context="detail_page")
        if not ok:
            log_error("detail_page", row.provider_detail_url, None, "failed after retries")
            continue

        sig = wait_for_page(driver, timeout=DETAIL_WAIT_TIMEOUT)
        if sig == "verification":
            sig = handle_verification(driver)
        if sig == "timeout":
            log_error("detail_page", row.provider_detail_url, None, "timed out waiting for content")
            continue
        if sig == "error":
            log_error("detail_page", row.provider_detail_url, None, "server error page detected")
            continue

        html = driver.page_source
        maybe_dump_debug_html(html, provider_id)

        try:
            fields = extract_detail_fields(html)
            row.fax = row.fax or fields["fax"]
            row.license_number = row.license_number or fields["license_number"]
            row.languages = row.languages or fields["languages"]
            row.office_hours = row.office_hours or fields["office_hours"]
        except Exception as e:
            log_error("detail_page_parse", row.provider_detail_url, None, str(e))

        row.detail_reviewed = "Y"

        print(f"Detail {i}/{len(to_visit)}: {row.provider_name} -- reviewed")

        if i % 5 == 0 or i == len(to_visit):
            save_all(providers)

        delay = random.uniform(*DETAIL_DELAY_RANGE)
        time.sleep(delay)


# ---------------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Sharp HealthTrio CA provider directory scraper")
    parser.add_argument("--start-url", default=ORIGINAL_TARGET_URL,
                         help="Entry search URL to start from (only used on a fresh run)")
    parser.add_argument("--fresh", action="store_true",
                         help="Ignore any existing progress and start over from startRow=0")
    parser.add_argument("--skip-details", action="store_true",
                         help="Skip the detail-page enrichment pass (fax/license/languages/hours)")
    parser.add_argument("--profile-dir", default=str(DEFAULT_PROFILE_DIR),
                         help="Chrome user-data-dir to launch with. Point this at a COPY of an "
                              "existing, already-used Chrome profile if the default throwaway "
                              "profile is triggering frequent verification. Chrome must not "
                              "already be running against that same directory.")
    args = parser.parse_args()
    profile_dir = Path(args.profile_dir).resolve()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log("=" * 60)
    log("Starting Sharp HealthTrio provider directory scrape")

    if args.fresh:
        for p in (STATE_PATH, XLSX_PATH, CSV_PATH):
            if p.exists():
                p.unlink()
        log("`--fresh` supplied: cleared previous progress files.")

    state = load_state()
    providers = load_existing_providers()

    resume_start_row = state.get("last_start_row", -10) + 10
    current_url = args.start_url if resume_start_row == 0 else state.get("last_page_url") or args.start_url
    start_row = resume_start_row
    total_count = state.get("total_count")

    log(f"Launching Chrome with profile directory: {profile_dir}")
    driver = build_driver(profile_dir)
    stop_flag = {"stop": False}

    try:
        page_num = start_row // 10 + 1
        while True:
            ok = safe_get(driver, current_url, context="results_page")
            if not ok:
                log_error("results_page", current_url, start_row, "failed after retries -- stopping")
                break

            sig = wait_for_page(driver)
            if sig == "verification":
                sig = handle_verification(driver)

            if sig == "timeout":
                log_error("results_page", current_url, start_row, "timed out waiting for content")
                print("Timed out waiting for the results page to load. Stopping so nothing is lost.")
                break
            if sig == "error":
                log_error("results_page", current_url, start_row, "server error page detected")
                print("A server error page was detected. Stopping so nothing is lost.")
                break
            if sig == "empty":
                print(f"No providers found at startRow={start_row} -- treating as end of results.")
                log(f"Empty results page at startRow={start_row}; stopping.")
                break

            actual_url = driver.current_url
            html = driver.page_source
            parsed = parse_results_page(html, actual_url, args.start_url)

            expected_first = start_row + 1
            if parsed["first_shown"] is not None and parsed["first_shown"] != expected_first:
                log(f"WARNING: stale/unexpected content at startRow={start_row} "
                    f"(page reports first item #{parsed['first_shown']}, expected #{expected_first}). "
                    f"Retrying this page once.")
                ok = safe_get(driver, current_url, context="results_page_retry")
                if ok:
                    wait_for_page(driver)
                    html = driver.page_source
                    actual_url = driver.current_url
                    parsed = parse_results_page(html, actual_url, args.start_url)

            new_count = 0
            for row in parsed["rows"]:
                key = row.dedup_key()
                if key not in providers:
                    providers[key] = row
                    new_count += 1

            if parsed["total_count"]:
                total_count = parsed["total_count"]
            approx_pages = math.ceil(total_count / 10) if total_count else "?"

            print(f"Page {page_num}/approximately {approx_pages}, startRow={start_row}: "
                  f"found {len(parsed['rows'])} providers ({new_count} new)")
            if parsed["errors"]:
                log(f"{len(parsed['errors'])} card(s) failed to parse on startRow={start_row}; see log above.")

            state.update({
                "last_start_row": start_row,
                "last_page_url": current_url,
                "total_count": total_count,
            })
            save_all(providers)
            save_state(state)
            print(f"Saved {len(providers)} unique providers")

            next_href = parsed["next_href"]
            if not next_href:
                print("No further 'Next Page' link found -- reached the end of results.")
                log("Reached end of results (no next link).")
                break

            current_url = next_href
            start_row += 10
            page_num += 1

            delay = random.uniform(*PAGE_DELAY_RANGE)
            print(f"Waiting {delay:.1f}s before the next page...")
            time.sleep(delay)

        if not args.skip_details:
            enrich_with_detail_pages(driver, providers, state, stop_flag)

    except KeyboardInterrupt:
        stop_flag["stop"] = True
        print("\nInterrupted (Ctrl+C). Progress already saved -- exiting safely.")
        log("Interrupted by user (Ctrl+C).")
    except Exception as e:
        log(f"FATAL ERROR: {e}\n{traceback.format_exc()}")
        print(f"A fatal error occurred: {e}\nSee {LOG_PATH} for details. Progress up to the last "
              f"successful save is intact.")
    finally:
        save_all(providers)
        save_state(state)
        print(f"\nFinal count: {len(providers)} unique providers saved to:\n"
              f"  {XLSX_PATH}\n  {CSV_PATH}\n  Log: {LOG_PATH}")
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()