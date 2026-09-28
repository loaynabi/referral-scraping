"""
Molina Healthcare California Medi-Cal PCP directory scraper -- REFERENCE /
DOCUMENTATION SCRIPT.

*** THIS SCRIPT CANNOT BE RUN AS-IS WITH `python scrape_molina.py`. ***
Molina's public provider-search site is a third-party Sapphire365 property
sitting behind Imperva bot-management (a `reese84` cookie challenge plus
per-request signing on every /api/providers/summary.json call). A plain
`requests`/`httpx` session -- even one that solves the initial reese84
cookie challenge -- gets HTTP 403 "Not Allowed" on every summary.json call
that it did not originate from the site's own obfuscated front-end JS. This
was confirmed empirically during this project: replaying the *exact* query
string (including the transaction_id/config_signature the page itself had
just used) via `fetch()` from the browser's own JS console, on the site's
own origin, with cookies included, STILL returned 403. Imperva is signing
each individual request client-side, not just gating on a static cookie.

The only reliable way to pull data found during this project was to drive
an actual, real browser session (e.g. via Claude's browser tools / Selenium
/ Playwright with a real Chromium instance -- NOT headless "requests-style"
automation) so the site's own JS makes and signs each request, and then
either:
  (a) intercept the XHR/fetch responses to /api/providers/summary.json, or
  (b) scrape the rendered result cards from the DOM after each search /
      "Next Page" click (what this project actually did -- see raw_*.txt
      and parse_raw.py in this folder).

Both approaches require a human-supervised or full-browser-automation
(Selenium/Playwright, non-headless or with stealth patches) session that
can pass the reese84 challenge. Neither urllib/requests/httpx scripts, nor
naive Selenium `driver.get()` + manual `requests.get(api_url)` combos, will
work -- the signing happens per-request inside the page's JS runtime.

------------------------------------------------------------------------
DOCUMENTED FLOW (what a Selenium/Playwright implementation should do)
------------------------------------------------------------------------

1. Launch a real (non-headless recommended) browser, navigate to:
       https://molina.sapphirethreesixtyfive.com/?ci=ca-medicaid
   Wait for the Imperva reese84 challenge to resolve (a few seconds) --
   the page title changes from "Sapphire365" to "Find a Healthcare
   Provider" once past it.

2. The "Plan/Program" defaults to Medi-Cal; the location box defaults to
   a rough Los Angeles location. Set the location explicitly:
     - Click the "City & State, County or Zip" field (or, on the results
       page, the compact one in the top nav) to open the Location panel.
     - Type a ZIP into the "Search by address, city, and zip code" field.
     - Click the matching suggestion, set "Within" (radius: 5/10/25/50/
       100 miles), click Apply.

3. To restrict to Primary Care Providers specifically, click the "Primary
   Care" quick-search chip on the homepage (or type into the main search
   box) and choose "Primary Care Provider (PCP)" from the autosuggest
   dropdown. This applies `search_specialty_id=980000011` and `is_pcp=Y`/
   `is_pcp=true` server-side -- confirmed via the network tab, this is
   the umbrella PCP category (Family Medicine, Internal Medicine, General
   Practice, Pediatrics, Geriatric Medicine, FQHCs, etc. all roll into it).

4. Remove the "Accepting New Patients" filter chip (X) if you want the
   full contracted PCP roster rather than only those currently open to
   new patients -- removing it noticeably increases the result count
   (e.g. for LA 90012/10mi: 1,148 -> 1,613 total PCPs).

5. Real API endpoint actually observed in the Network tab, for reference:

   GET https://molina.sapphirethreesixtyfive.com/api/providers/summary.json
     ?network_id=13
     &geo_location={lat},{lng}
     &locale=en&data_language=en
     &search_specialty_id=980000011      # PCP umbrella category
     &page={n}
     &radius={5|10|25|50|100}
     &limit=10                            # UI hard-codes 10/page; larger
                                           # limit values (e.g. 100) get
                                           # a 403, confirming server-side
                                           # validation tied to the signed
                                           # request, not just a query param
     &sort=...
     &is_pcp=true                         # or contract_accepting_new_patients=Y&is_pcp=Y
     &transaction_id={uuid}               # single-use, tied to the signed
                                           # request -- cannot be reused
     &account_id=68
     &ci=ca-medicaid
     &config_signature={opaque}

   Response JSON shape (per page):
     {
       "_meta": {"counts": {"total": N}, "pages": {"total": P, "current": c,
                 "next": c+1, "previous": c-1}},
       "providers": [
         {
           "name", "first_name", "last_name", "middle_name",
           "primary_specialty", "additional_specialties": [...],
           "identifiers": [{"type_code": "NPI", "value": "..."}, ...],
           "phone", "fax", "addr_line1", "addr_line2", "city", "state",
           "postal_code", "location_name", "medical_group",
           "accepting_new_patients": "Y"/"N",
           "languages": [...], "is_pcp": true/false,
           "latitude", "longitude", "distance", ...
         }, ...
       ]
     }

   `is_pcp` (boolean/Y-N) is exactly the field to filter on if you ever DO
   get raw JSON (e.g. by intercepting the real browser's network traffic
   with a tool like Claude's `read_network_requests`, mitmproxy, or the
   Chrome DevTools Protocol) -- no need for a separate specialty allowlist.

6. Pagination: click the "Next Page" button (`button[aria-label="Go to
   next page"]` / accessible name "Go to next page") and re-read the
   result list or re-capture the network response each time. The button's
   underlying DOM node is re-created on every page load, so re-locate it
   fresh each time rather than caching a stale element handle.

7. If a search+filter combination narrows the total below 1,000 results,
   the site itself exposes a native "Download" button (Download Search
   Results) next to "Share" beneath the results list, which presumably
   exports the full filtered set in one shot. This project deliberately
   did NOT click it (downloading a third-party file requires separate,
   explicit user sign-off per this agent's operating rules) -- a future
   run with that permission granted could use it to avoid manual paging
   entirely for narrower geographic/specialty slices (e.g. per-county
   searches, which frequently land under the 1,000-result cap).

------------------------------------------------------------------------
WHAT THIS PROJECT ACTUALLY DID
------------------------------------------------------------------------
Given the anti-automation constraints above, actual data collection for
this project was done via a live, tool-driven browser session (Claude's
`Claude_Browser` MCP tools), searching "Primary Care Provider (PCP)" (with
the "Accepting New Patients" filter removed) for five CA metro ZIP codes
(Los Angeles 90012, San Diego 92101, Sacramento 95814, Fresno 93701, San
Jose 95112) at the default 10-mile radius, capturing the first 1-3 pages
of results per metro (`get_page_text` on the rendered DOM, since raw
network-response bodies were being truncated by the capture tool at very
large sizes). Those DOM captures were saved to `raw_<metro>_<zip>_page*.txt`
in this folder, then parsed by `parse_raw.py` into `checkpoint.json`, then
turned into the final `Molina_HealthCare_Providers.xlsx` by `build_excel.py`.

This is a PARTIAL SAMPLE, not an exhaustive statewide pull -- see this
project's final chat summary for exact row counts and geographic coverage
achieved (Molina's PCP network size varies enormously by metro: e.g.
Los Angeles and San Diego each have 1,000+ in-network PCPs within 10 miles
of the sampled ZIP, while Fresno had only 14 and San Jose only 3, likely
because Molina is not the dominant Medi-Cal managed-care plan in those
counties). Re-running with more turns/time, and paging through fully
(or splitting each metro into multiple sub-1,000-result slices to unlock
the native Download button, with explicit user permission to download
that file), would be the way to reach full coverage.
"""

raise SystemExit(
    "This is a documentation/reference script, not a runnable scraper -- "
    "see the module docstring above for why, and for the flow a real "
    "browser-automation implementation should follow."
)
