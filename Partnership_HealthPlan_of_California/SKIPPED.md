# SKIPPED — Partnership HealthPlan of California PCP Directory Scrape

**Date attempted:** 2026-09-09
**Target:** https://providerdirectory.partnershiphp.org/Provider/BasicSearch/

## Outcome

No provider data was collected. This was a **hard blocker**, not a data-quality or
scope issue — no rows were fabricated or approximated in place of real results.

## What happened

Every attempt to load the search page (via the Claude Browser MCP tool, a real
JS-executing browser) returned:

```
partnershiphp.org | 522: Connection timed out
```

HTTP 522 is a Cloudflare-specific error meaning **Cloudflare's edge could not
establish a connection to the origin web server at all** — it is emitted before
any app content, WAF challenge, or "Just a moment..." JS challenge page is served.
This is materially different from the bot-detection/JS-challenge scenario
described in the prior investigation (which *does* typically resolve for a real
browser session given time). A 522 indicates the origin server itself is down,
overloaded, unreachable from Cloudflare's network, or blackholing Cloudflare's
requests — none of which a client-side wait or slower pacing can fix.

## Steps taken before giving up

1. Navigated to the BasicSearch URL — got 522 immediately.
2. Waited (5s), retried — 522 again.
3. Repeated wait-then-navigate cycles (multiple rounds of 10s waits) totaling
   roughly 2 minutes of elapsed time, with fresh full page loads (not rapid
   clicking/polling) between each — the response was 522 every single time,
   with no variation (no challenge page, no partial content, no redirect to a
   real form).
4. Confirmed no CAPTCHA or "Just a moment..." interstitial was ever shown —
   only the flat Cloudflare 522 error page — so there was nothing to "wait out"
   in the way a JS challenge would resolve.

## No data fabricated

Per instructions, since the site could not be reached at all, no checkpoint
file, no Excel file, and no provider rows were produced. Writing plausible-
looking sample data would misrepresent this as real public directory data,
which was explicitly disallowed.

## Recommendation for re-run

- Retry later — a 522 is usually transient (origin server restart, deploy,
  or temporary overload) rather than a permanent block. Check
  https://providerdirectory.partnershiphp.org/ status again in a few hours,
  or the next business day.
- If 522s persist across multiple days, contact Partnership HealthPlan of
  California provider relations directly, or use their published/downloadable
  PCP directory PDF/CSV (many Medi-Cal plans publish a static directory file
  alongside the interactive search tool) as an alternative public data source.
- When the site is reachable again, use the scraper in `scrape_phc.py` in this
  same folder as the starting point — it documents the known form fields,
  selectors, and a slow/human-paced automation strategy for the Cloudflare
  bot protection that the prior investigation *did* encounter (a separate,
  real JS-challenge concern, distinct from today's 522).
