"""Interactive Sharp Health Plan (HealthTrio) California provider scraper.

Run from Command Prompt or PowerShell:
    py healthtrio_scraper.py

The browser is intentionally visible. If Cloudflare appears, complete it
manually in the browser and press Enter in the terminal when the results page
is visible. The script never attempts to bypass the verification.
"""

from __future__ import annotations

import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright


START_URL = (
    "https://sharp.healthtrioconnect.com/public-app/consumer/provdir/"
    "search.page?isEntry=entry&startRow=0&sortType=NAMEASC&provType=P&addressState=ca"
)
PAGE_DELAY_SECONDS = 8
PROVIDER_DELAY_SECONDS = 1.5
OUTPUT = Path(__file__).with_name("output") / "sharp_healthtrio_providers.xlsx"
COLUMNS = [
    "Name", "Specialty", "NPI", "Phone", "Fax", "Address", "City", "State",
    "ZIP", "Accepting New Patients", "Accepting New Medicare Patients",
    "Medical Group", "Practice", "Provider Detail URL", "Source", "Scraped At",
]


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def verification_visible(page) -> bool:
    title = page.title().lower()
    body = clean(page.locator("body").inner_text()).lower()
    return "just a moment" in title or "verify you are human" in body or "cloudflare" in body


def parse_record(item, page_url: str, scraped_at: str) -> dict[str, str]:
    text = [clean(value) for value in item.inner_text().splitlines() if clean(value)]
    detail = item.locator("a[href*='providerDetail.page']").first
    name = clean(detail.inner_text()) if detail.count() else (text[0] if text else "")
    joined = " | ".join(text)
    city_line_index = next(
        (index for index, value in enumerate(text)
         if re.search(r",\s*[A-Z]{2}\s+\d{5}(?:-\d{4})?$", value)), -1
    )
    address_match = re.search(
        r"^(.+?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$",
        text[city_line_index] if city_line_index >= 0 else "",
    )
    phone_match = re.search(r"Phone:\s*([0-9().+\-\s]+)", joined, re.I)
    fax_match = re.search(r"Fax:\s*([0-9().+\-\s]+)", joined, re.I)
    npi_match = re.search(r"Provider NPI:\s*([0-9]+)", joined, re.I)
    href = detail.get_attribute("href") if detail.count() else ""
    address = text[city_line_index - 1] if address_match and city_line_index > 0 else ""
    accepting_new = bool(re.search(r"Accepting New Patients", joined, re.I)) and not bool(
        re.search(r"Not Accepting New Patients", joined, re.I)
    )
    return {
        "Name": name,
        "Specialty": next((x for x in text if x.lower() not in {name.lower()} and "provider npi" not in x.lower()
                           and "accepting" not in x.lower() and "accessibility" not in x.lower()
                           and "phone:" not in x.lower() and "fax:" not in x.lower()), ""),
        "NPI": npi_match.group(1) if npi_match else "",
        "Phone": clean(phone_match.group(1)) if phone_match else "",
        "Fax": clean(fax_match.group(1)) if fax_match else "",
        "Address": address,
        "City": address_match.group(1).split(",")[-1].strip() if address_match else "",
        "State": address_match.group(2) if address_match else "CA",
        "ZIP": address_match.group(3) if address_match else "",
        "Accepting New Patients": "Yes" if accepting_new else "No",
        "Accepting New Medicare Patients": "Yes" if re.search(r"Accepting New Medicare Patients", joined, re.I)
        and not re.search(r"Not Accepting New Medicare Patients", joined, re.I) else "No",
        "Medical Group": text[-5] if len(text) >= 5 else "",
        "Practice": text[-4] if len(text) >= 4 else "",
        "Provider Detail URL": urljoin(page_url, href) if href else "",
        "Source": START_URL,
        "Scraped At": scraped_at,
    }


def save(rows: list[dict[str, str]]) -> None:
    OUTPUT.parent.mkdir(exist_ok=True)
    pd.DataFrame(rows, columns=COLUMNS).drop_duplicates(
        subset=["Name", "Provider Detail URL"]
    ).to_excel(OUTPUT, index=False, engine="openpyxl")


def main() -> None:
    scraped_at = datetime.now(timezone.utc).isoformat()
    rows: list[dict[str, str]] = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(START_URL, wait_until="domcontentloaded", timeout=90_000)
        print("HealthTrio opened in a visible browser.")
        print("Complete Cloudflare if shown; waiting for provider results...", flush=True)
        for attempt in range(120):
            if page.locator("a[href*='providerDetail.page']").count():
                break
            page.wait_for_timeout(1000)
        else:
            raise RuntimeError(
                "Provider results did not appear after two minutes. "
                "Complete verification in the visible browser and restart the script."
            )
        for page_number in range(1, 200):
            page.wait_for_timeout(1000)
            try:
                page.wait_for_selector("a[href*='providerDetail.page']", timeout=20_000)
            except PlaywrightTimeoutError:
                print(f"Page {page_number}: no provider results found; stopping", flush=True)
                break
            detail_links = page.locator("a[href*='providerDetail.page']")
            hrefs = []
            for index in range(detail_links.count()):
                href = detail_links.nth(index).get_attribute("href")
                if href and href not in hrefs:
                    hrefs.append(href)
            count = len(hrefs)
            print(f"Page {page_number}: found {count} provider cards", flush=True)
            for href in hrefs:
                detail_link = page.locator(f"a[href='{href}']").first
                item = detail_link.locator("xpath=ancestor::li[1]")
                rows.append(parse_record(item, page.url, scraped_at))
                time.sleep(PROVIDER_DELAY_SECONDS)
            save(rows)
            next_link = page.get_by_role("link", name="Next Page").first
            if not next_link.count():
                break
            href = next_link.get_attribute("href")
            if not href:
                break
            print(f"Waiting {PAGE_DELAY_SECONDS}s before opening the next page...", flush=True)
            time.sleep(PAGE_DELAY_SECONDS)
            page.goto(urljoin(page.url, href), wait_until="domcontentloaded", timeout=90_000)
            time.sleep(PAGE_DELAY_SECONDS)
            for _ in range(180):
                if page.locator("a[href*='providerDetail.page']").count():
                    break
                if verification_visible(page):
                    print("Verification detected; complete it in the browser...", flush=True)
                page.wait_for_timeout(1000)
            else:
                print("Timed out waiting for the next page; stopping.", flush=True)
                break
        browser.close()
    save(rows)
    print(f"Complete: {len(rows)} rows saved to {OUTPUT}")


if __name__ == "__main__":
    main()
