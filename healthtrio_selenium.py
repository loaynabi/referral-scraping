"""Slow, visible Selenium scraper for the Sharp Health Plan HealthTrio directory.

Run:
    py healthtrio_selenium.py

Complete any Cloudflare check in the visible Chrome window. The scraper does
not bypass verification and does not open provider pages concurrently.
"""

from __future__ import annotations

import argparse
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait


START_URL = (
    "https://sharp.healthtrioconnect.com/public-app/consumer/provdir/"
    "search.page?isEntry=entry&startRow=0&sortType=NAMEASC&provType=P&addressState=ca"
)
OUTPUT = Path(__file__).with_name("output") / "sharp_healthtrio_providers_selenium.xlsx"
DEFAULT_PROFILE = Path(__file__).with_name("chrome-healthtrio-profile")
COLUMNS = [
    "Name", "Specialty", "NPI", "Phone", "Fax", "Address", "City", "State",
    "ZIP", "Accepting New Patients", "Accepting New Medicare Patients",
    "Medical Group", "Practice", "Provider Detail URL", "Source", "Scraped At",
]


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def wait_for_results(
    driver: webdriver.Chrome, expected_start_row: int, timeout: int = 240
) -> bool:
    try:
        WebDriverWait(driver, timeout).until(
            lambda browser: (
                len(browser.find_elements(By.CSS_SELECTOR, "a[href*='providerDetail.page']")) > 0
                and f"startRow={expected_start_row}" in browser.current_url
            )
        )
        return True
    except TimeoutException:
        return False


def parse_record(item, page_url: str, scraped_at: str) -> dict[str, str]:
    text = [clean(value) for value in item.text.splitlines() if clean(value)]
    links = item.find_elements(By.CSS_SELECTOR, "a[href*='providerDetail.page']")
    name = clean(links[0].text) if links else (text[0] if text else "")
    joined = " | ".join(text)
    city_line = next((value for value in text if re.search(
        r",\s*[A-Z]{2}\s+\d{5}(?:-\d{4})?$", value)), "")
    location = re.search(r"^(.+?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line)
    phone = re.search(r"Phone:\s*([0-9().+\-\s]+)", joined, re.I)
    fax = re.search(r"Fax:\s*([0-9().+\-\s]+)", joined, re.I)
    npi = re.search(r"Provider NPI:\s*([0-9]+)", joined, re.I)
    href = links[0].get_attribute("href") if links else ""
    city_index = text.index(city_line) if city_line else -1
    address = text[city_index - 1] if city_index > 0 else ""
    accepting = "Yes" if re.search(r"Accepting New Patients", joined, re.I) and not re.search(
        r"Not Accepting New Patients", joined, re.I) else "No"
    accepting_medicare = "Yes" if re.search(
        r"Accepting New Medicare Patients", joined, re.I
    ) and not re.search(r"Not Accepting New Medicare Patients", joined, re.I) else "No"
    excluded = {name.lower(), "accepting new patients", "not accepting new patients"}
    specialty = next((value for value in text if value.lower() not in excluded
                      and "provider npi" not in value.lower()
                      and "phone:" not in value.lower()
                      and "fax:" not in value.lower()
                      and "accepting" not in value.lower()), "")
    return {
        "Name": name,
        "Specialty": specialty,
        "NPI": npi.group(1) if npi else "",
        "Phone": clean(phone.group(1)) if phone else "",
        "Fax": clean(fax.group(1)) if fax else "",
        "Address": address,
        "City": location.group(1).split(",")[-1].strip() if location else "",
        "State": location.group(2) if location else "CA",
        "ZIP": location.group(3) if location else "",
        "Accepting New Patients": accepting,
        "Accepting New Medicare Patients": accepting_medicare,
        "Medical Group": text[-5] if len(text) >= 5 else "",
        "Practice": text[-4] if len(text) >= 4 else "",
        "Provider Detail URL": urljoin(page_url, href) if href else "",
        "Source": START_URL,
        "Scraped At": scraped_at,
    }


def save(rows: list[dict[str, str]]) -> None:
    OUTPUT.parent.mkdir(exist_ok=True)
    frame = pd.DataFrame(rows, columns=COLUMNS).drop_duplicates(
        subset=["Name", "Provider Detail URL"]
    )
    try:
        frame.to_excel(OUTPUT, index=False, engine="openpyxl")
    except PermissionError:
        fallback = OUTPUT.with_name("sharp_healthtrio_providers_selenium_progress.xlsx")
        frame.to_excel(fallback, index=False, engine="openpyxl")
        print(f"Workbook is open; saved {len(frame)} rows to {fallback}", flush=True)
        return
    print(f"Saved {len(frame)} rows to {OUTPUT}", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile-dir", type=Path, default=DEFAULT_PROFILE)
    parser.add_argument("--page-delay", type=float, default=12.0)
    parser.add_argument("--provider-delay", type=float, default=3.0)
    parser.add_argument("--start-url", default=START_URL)
    parser.add_argument("--max-pages", type=int, default=200)
    args = parser.parse_args()
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={args.profile_dir.resolve()}")
    driver = webdriver.Chrome(options=options)
    rows: list[dict[str, str]] = []
    scraped_at = datetime.now(timezone.utc).isoformat()
    try:
        driver.get(args.start_url)
        print("Visible Chrome opened. Complete Cloudflare manually if shown.", flush=True)
        start_match = re.search(r"[?&]startRow=(\d+)", args.start_url)
        start_row = int(start_match.group(1)) if start_match else 0
        if not wait_for_results(driver, start_row):
            raise RuntimeError("Provider results did not appear within four minutes.")
        for page_number in range(1, args.max_pages + 1):
            time.sleep(args.page_delay)
            current_match = re.search(r"[?&]startRow=(\d+)", driver.current_url)
            current_row = int(current_match.group(1)) if current_match else start_row
            if not wait_for_results(driver, current_row):
                print(f"Page {page_number}: waiting for verification/results timed out.", flush=True)
                break
            links = driver.find_elements(By.CSS_SELECTOR, "a[href*='providerDetail.page']")
            unique_links = {}
            for link in links:
                href = link.get_attribute("href")
                if href:
                    unique_links.setdefault(href, link)
            print(f"Page {page_number}: {len(unique_links)} providers", flush=True)
            for href, link in unique_links.items():
                item = link.find_element(By.XPATH, "./ancestor::li[1]")
                rows.append(parse_record(item, driver.current_url, scraped_at))
                time.sleep(args.provider_delay)
            save(rows)
            next_links = driver.find_elements(By.CSS_SELECTOR, "a[aria-label='Next Page']")
            if not next_links:
                break
            next_href = next_links[0].get_attribute("href")
            if not next_href:
                break
            next_match = re.search(r"[?&]startRow=(\d+)", next_href)
            next_row = int(next_match.group(1)) if next_match else current_row + 10
            print(f"Waiting {args.page_delay}s before page {page_number + 1}...", flush=True)
            time.sleep(args.page_delay)
            driver.get(urljoin(driver.current_url, next_href))
            if not wait_for_results(driver, next_row):
                print(
                    f"Page {page_number + 1}: waiting for verification/results timed out.",
                    flush=True,
                )
                break
        save(rows)
        print(f"Complete: {len(rows)} rows", flush=True)
    except WebDriverException as error:
        print(f"Browser error: {error}", flush=True)
        raise
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
