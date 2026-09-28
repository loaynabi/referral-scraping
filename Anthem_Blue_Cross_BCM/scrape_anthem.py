"""
Reusable starting point for a Selenium-based scrape of Anthem Blue Cross
California's public "Find Care" directory, filtered to the Medi-Cal HMO
network / Primary Care category.

STATUS: This script documents and encodes the flow and CSS/DOM selectors
that were confirmed to work against https://www.anthem.com/ca/find-care/
during manual browser-automation testing (Sept 2026). It has NOT been run
end-to-end unattended -- see "Known fragility" below before relying on it
for a large unattended run. Treat it as a documented starting point, not a
finished production scraper.

Background / why this exists
-----------------------------
Anthem's Find Care tool is a client-side Angular app sitting behind Akamai
Bot Manager. There is no clean, directly-callable JSON API discoverable from
the network tab -- search results are fetched via obfuscated Akamai sensor
POSTs, and the actual provider list is only ever available by driving the
real UI. This script drives the real UI with Selenium (a real Chrome
browser, not headless requests) and scrapes the rendered
`app-fc-provider-card-cmp` elements.

Flow (confirmed working, in order)
-----------------------------------
1. GET https://www.anthem.com/ca/find-care/?cnslocale=en_US_ca
2. Account type step:
     - radio input#accountTypeMedicaid -> check it (usually pre-selected)
     - button#... "Continue" if the account-type radios were not already set
3. Click button#open-plan-search-modal ("Basic search as a guest").
   This opens a slide-in side panel: div.pf-side-panel.fcr-plan-slider
   NOTE: on this environment the panel sometimes renders far outside the
   visible viewport (a real rendering bug, not just scroll position) --
   coordinate/pixel clicks on it are unreliable. Prefer JS-level
   interaction (see selectors below) over Selenium's normal
   element.click() + ActionChains where possible.
4. Inside the panel there are 4 "auto-suggest" combo fields, in this order:
     0. #uxd-auto-suggest-cmp-0-q-field  -- "Select the type of plan or network"
        -> type/select "Medical Plan or Network ..."
        -> option labels render as #uxd-auto-suggest-cmp-0-ddlLabel0, ddlLabel1, ...
     1. #uxd-auto-suggest-cmp-1-q-field  -- "Select the state ..."
        -> type/select "California" (was ddlLabel4 in the unfiltered 55-state list)
     2. #uxd-auto-suggest-cmp-2-q-field  -- "Select how you get health insurance"
        -> type/select "Medicaid" (was ddlLabel2 in the unfiltered list)
     3. #uxd-auto-suggest-cmp-3-q-field  -- "Select a plan or network"
        -> only one option once the above are set: "Medi-Cal HMO" (ddlLabel0)
   Each field is a custom Angular combobox, not a native <select>. To open
   the option list we found two techniques that worked in different runs
   (this component was noticeably flaky / non-deterministic across runs --
   see "Known fragility"):
     (a) set the input's value via the native setter + dispatch an 'input'
         event (bypasses React/Angular's normal onChange but is picked up
         by this component's ngModel binding), e.g.:
             the_input.dispatchEvent(new Event('input', {bubbles: true}))
     (b) call `the_input.click()` directly (opens the dropdown as a
         side-effect of focus/click handling)
   After the labels render, click the matching `label.psLabel` element
   (id pattern `#uxd-auto-suggest-cmp-{N}-ddlLabel{index}`).
5. Click button#btn-plan-continue to submit the plan-selection panel.
6. Location step: click .btn-update-location (opens a "Search Location"
   panel), then:
     - select radio#selectZipCityCounty
     - type the ZIP into input#zipCityCounty
     - click button#btn-assign-pcp-submit ("Continue")
7. Category step: on the resulting "Search by Care Provider" screen, click
   the quick-link anchor whose visible label is "Primary Care"
   (aria-label="select for Primary Care based provider search",
   id pattern like #lnk-providerType-0 -- id is NOT stable, match by the
   inner <span class="fcr-iconCardLabel"> text instead).
8. This navigates to the results view. Wait for
   `app-fc-provider-card-cmp` elements (20 per page) or, if the page comes
   back empty, for the `.motif-spinner-heart` loading spinner to disappear
   (in testing this occasionally spun indefinitely -- see fragility notes).

Result card scraping
---------------------
Each `app-fc-provider-card-cmp` element's `.text` (innerText) is a newline-
delimited block shaped like:

    In-Network
    <Name>, <credentials>
    <Specialty>
    <Street address>
    <City>, <ST> <ZIP> (<distance> mi)
    <Phone>

Parse it line-by-line (see `parse_card_text` below, shared logic with
build_output.py in this folder).

Pagination
----------
- Pagination lives in `app-fcr-pagination`, with a `<ul id="pagination">`
  of `<button id="lnk-pf-pagination-N">` page buttons plus
  `#lnk-pf-pagination-next` / `#lnk-pf-pagination-previous` /
  `#lnk-pf-pagination-first` / `#lnk-pf-pagination-last`.
- The summary text "Showing 1-20 of 2000 results within 20 miles of 90012"
  lives in `.fcr-pagination-summary`.
- To go to the next page: click `#lnk-pf-pagination-next`, then poll until
  the first card's text changes (Angular re-renders in place, so a naive
  "wait for element to exist" is not enough -- compare content).

Known fragility (encountered during manual testing, budget this into any
unattended run)
-----------------------------------------------------------------------
1. After roughly 5-8 consecutive rapid page-to-page clicks, the results
   list silently emptied out (`app-fc-provider-list-cmp` disappeared from
   the DOM entirely, no error shown, no failed network request visible)
   and never recovered without reloading the whole flow from step 1. This
   looks like an Akamai Bot Manager soft mitigation triggered by
   non-human click cadence. Mitigation: add randomized human-like delays
   (1-3s) between page clicks, consider scrolling/mouse-move jitter
   between actions, and checkpoint frequently so a stall only costs a few
   pages.
2. The plan-selection auto-suggest widgets (step 4 above) were
   inconsistently interactive across runs: sometimes `.click()` /
   dispatching an `input` event opened the option list immediately,
   other times identical code produced no dropdown at all with no
   apparent difference in DOM state. Root cause not confirmed (possibly
   Angular zone.js change-detection timing, possibly a bot-detection
   countermeasure). Mitigation: retry the open-dropdown step a few times
   with waits before giving up, and consider using a real
   Selenium ActionChains click (native OS-level click) instead of a
   synthetic JS click, since this may behave more consistently than
   JS-dispatched events for this specific widget.
3. The plan-selection side panel has, in some renders, a CSS bug that
   positions it mostly or entirely outside the visible viewport
   (`getBoundingClientRect().x` computed at ~1265px in a 1280px-wide
   viewport, i.e. only ~15px visible) even though it is not otherwise
   hidden. Native pixel-coordinate clicks fail in that state; only
   JS-level `element.click()` / event dispatch reaches it.
4. This whole flow is Angular-router-based; a plain `driver.get()` /
   history navigation back to the criteria URL usually drops all
   in-app state and restarts at the account-type screen. Do not expect a
   URL to durably encode search state -- treat this as a stateful wizard,
   not a set of bookmarkable pages.

Given (1)-(3), this script is written to checkpoint provider rows to a
JSON file after every page, so an interrupted run loses at most one page
of results.

Usage
-----
    pip install selenium webdriver-manager
    python scrape_anthem.py --zip 90012 --max-pages 100 \
        --out raw_selenium_90012.json

This writes an array-of-arrays JSON file in the same shape as the
raw/*.json files already in this folder, so it can be fed straight into
build_output.py's `load_all_rows()`.
"""
import argparse
import json
import re
import time
import random
import sys

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        TimeoutException,
        NoSuchElementException,
        ElementNotInteractableException,
    )
except ImportError:
    print("This script requires selenium. Install with: pip install selenium")
    sys.exit(1)

BASE_URL = "https://www.anthem.com/ca/find-care/?cnslocale=en_US_ca"


def human_wait(a=0.8, b=2.2):
    time.sleep(random.uniform(a, b))


def js_set_value_and_input(driver, element, value):
    driver.execute_script(
        """
        const el = arguments[0], value = arguments[1];
        const proto = Object.getPrototypeOf(el);
        const desc = Object.getOwnPropertyDescriptor(proto, 'value');
        desc.set.call(el, value);
        el.dispatchEvent(new Event('input', {bubbles: true}));
        """,
        element,
        value,
    )


def js_click(driver, element):
    driver.execute_script("arguments[0].click();", element)


def open_dropdown_and_select(driver, wait, field_index, type_text, option_id):
    """Open one of the 4 auto-suggest combo fields and click a specific option.

    field_index: 0=plan type, 1=state, 2=insurance type, 3=plan/network
    type_text: text to type into the field (use '' if the option list is
        already short enough to render unfiltered, e.g. field 3)
    option_id: element id of the `label.psLabel` to click once visible,
        e.g. 'uxd-auto-suggest-cmp-1-ddlLabel4' for California.
    """
    field = driver.find_element(By.ID, f"uxd-auto-suggest-cmp-{field_index}-q-field")

    for attempt in range(3):
        try:
            js_set_value_and_input(driver, field, type_text)
        except Exception:
            pass
        try:
            js_click(driver, field)
        except Exception:
            pass
        human_wait(0.8, 1.4)
        try:
            option = wait.until(EC.presence_of_element_located((By.ID, option_id)))
            js_click(driver, option)
            human_wait(0.6, 1.1)
            return True
        except TimeoutException:
            continue
    return False


def run_search_setup(driver, wait, zip_code):
    driver.get(BASE_URL)
    human_wait(2, 3)

    # Account type: Medicaid
    try:
        radio = driver.find_element(By.ID, "accountTypeMedicaid")
        if not radio.is_selected():
            js_click(driver, radio)
            human_wait()
    except NoSuchElementException:
        pass

    # Open "Basic search as a guest" modal
    open_btn = wait.until(EC.presence_of_element_located((By.ID, "open-plan-search-modal")))
    js_click(driver, open_btn)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".pf-side-panel.fcr-plan-slider")))
    human_wait()

    ok = True
    ok &= open_dropdown_and_select(driver, wait, 0, "Medical", "uxd-auto-suggest-cmp-0-ddlLabel0")
    ok &= open_dropdown_and_select(driver, wait, 1, "California", "uxd-auto-suggest-cmp-1-ddlLabel4")
    ok &= open_dropdown_and_select(driver, wait, 2, "", "uxd-auto-suggest-cmp-2-ddlLabel2")
    ok &= open_dropdown_and_select(driver, wait, 3, "", "uxd-auto-suggest-cmp-3-ddlLabel0")
    if not ok:
        raise RuntimeError("Failed to complete plan-selection panel (see Known fragility #2 in module docstring)")

    continue_btn = driver.find_element(By.ID, "btn-plan-continue")
    js_click(driver, continue_btn)
    human_wait(1.5, 2.5)

    # Location
    update_loc = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".btn-update-location")))
    js_click(driver, update_loc)
    human_wait()
    zip_input = wait.until(EC.presence_of_element_located((By.ID, "zipCityCounty")))
    js_set_value_and_input(driver, zip_input, zip_code)
    human_wait(0.6, 1.0)
    radio = driver.find_element(By.ID, "selectZipCityCounty")
    js_click(driver, radio)
    human_wait(0.4, 0.7)
    submit_btn = driver.find_element(By.ID, "btn-assign-pcp-submit")
    js_click(driver, submit_btn)
    human_wait(1.5, 2.5)

    # Category: Primary Care (match by visible label text, id is not stable)
    primary_care_link = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//span[contains(@class,'fcr-iconCardLabel') and normalize-space(text())='Primary Care']/ancestor::a")
        )
    )
    js_click(driver, primary_care_link)
    human_wait(2.5, 4)

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-fc-provider-card-cmp")))


CARD_TAG = "app-fc-provider-card-cmp"


def extract_cards(driver):
    cards = driver.find_elements(By.TAG_NAME, CARD_TAG)
    return [c.text.strip() for c in cards if c.text.strip()]


def go_next_page(driver, wait, prev_first_card_text, timeout=10):
    try:
        next_btn = driver.find_element(By.ID, "lnk-pf-pagination-next")
    except NoSuchElementException:
        return False
    parent_classes = next_btn.find_element(By.XPATH, "..").get_attribute("class") or ""
    if "disabled" in parent_classes:
        return False
    js_click(driver, next_btn)
    start = time.time()
    while time.time() - start < timeout:
        cards = extract_cards(driver)
        if cards and cards[0] != prev_first_card_text:
            return True
        time.sleep(0.4)
    return False


def parse_card_text(card_text, source_zip):
    """Mirrors build_output.py's parse_card() -- kept in sync manually."""
    lines = [l.strip() for l in card_text.split("\n") if l.strip()]
    if not lines:
        return None
    idx = 0
    network_status = None
    if lines[0] in ("In-Network", "Out-of-Network", "In-Network*"):
        network_status = lines[0]
        idx = 1
    name = lines[idx] if idx < len(lines) else ""
    specialty = lines[idx + 1] if idx + 1 < len(lines) else ""
    street = lines[idx + 2] if idx + 2 < len(lines) else ""
    csz_line = lines[idx + 3] if idx + 3 < len(lines) else ""
    phone = lines[idx + 4] if idx + 4 < len(lines) else ""

    m = re.match(
        r"^(?P<city>.+),\s*(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})(?:-\d{4})?\s*(?:\((?P<distance>[\d.]+)\s*mi\))?\s*$",
        csz_line,
    )
    city = state = zip_code = distance = ""
    if m:
        city, state, zip_code, distance = (
            m.group("city"),
            m.group("state"),
            m.group("zip"),
            m.group("distance") or "",
        )

    return {
        "Provider Type": "Primary Care",
        "Name": name,
        "Specialty": specialty,
        "Network Status": network_status or "",
        "Street Address": street,
        "City": city,
        "State": state,
        "Zip": zip_code,
        "Distance (mi)": distance,
        "Phone": phone,
        "Search Zip": source_zip,
        "Source": "Anthem Blue Cross CA Find Care - Medi-Cal HMO",
    }


def scrape_zip(driver, wait, zip_code, max_pages, checkpoint_path):
    run_search_setup(driver, wait, zip_code)

    all_pages_raw = []  # array-of-arrays of raw card text, for build_output.py compatibility
    page_num = 1
    while page_num <= max_pages:
        cards = extract_cards(driver)
        if not cards:
            print(f"[{zip_code}] page {page_num}: no cards found, stopping")
            break
        all_pages_raw.append(cards)
        print(f"[{zip_code}] page {page_num}: {len(cards)} cards")

        # Checkpoint every page
        with open(checkpoint_path, "w", encoding="utf-8") as f:
            json.dump(all_pages_raw, f, indent=2, ensure_ascii=False)

        prev_first = cards[0]
        advanced = go_next_page(driver, wait, prev_first)
        if not advanced:
            print(f"[{zip_code}] could not advance past page {page_num} "
                  f"(end of results, or hit the Known-fragility #1 stall)")
            break
        human_wait(1.0, 2.5)  # human-like pacing between pages
        page_num += 1

    return all_pages_raw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--zip", required=True, help="ZIP code to search, e.g. 90012")
    ap.add_argument("--max-pages", type=int, default=100)
    ap.add_argument("--out", default=None, help="Output JSON path (array-of-arrays of raw card text)")
    ap.add_argument("--headless", action="store_true")
    args = ap.parse_args()

    out_path = args.out or f"raw_selenium_{args.zip}.json"

    options = webdriver.ChromeOptions()
    if args.headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1280,900")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)

    try:
        scrape_zip(driver, wait, args.zip, args.max_pages, out_path)
    finally:
        driver.quit()

    print(f"Done. Raw pages written to {out_path}")
    print("Feed this into build_output.py's load_all_rows() (or copy it into raw/) to merge into the Excel output.")


if __name__ == "__main__":
    main()
