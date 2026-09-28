Source note for Mountain_Valley_Health_Plan_Providers.xlsx
============================================================

This spreadsheet is a duplicate of:
  ..\Health_Plan_of_San_Joaquin\Health_Plan_of_San_Joaquin_Providers.xlsx

Both "Health Plan of San Joaquin" (HPSJ) and "Mountain Valley Health Plan" (MVHP)
are covered by the SAME combined Medi-Cal provider directory PDF:

  HPSJ-MVHP-Medi-Cal-Provider-Directory-September-2026.pdf

This PDF explicitly states it serves both "Health Plan of San Joaquin" and
"Mountain Valley Health Plan" members, and shares one identical provider
network/directory. There is no separate MVHP-only source document, so this
file is provided as-is for the MVHP entry in the referral-coordination
project, duplicated from the HPSJ extraction rather than re-scraped.

Extraction details (Primary Care Doctors section only):
- Source PDF pages 32-377 (0-indexed), bounded by the "Primary Care Doctors"
  section banner (page 32) and the "Obstetrics/Gynecologists Doctors" section
  banner (page 378), which marks the start of the next section.
- Specialties included: Family Medicine, Internal Medicine, Pediatrics,
  General Practice (as they appear under the "Primary Care Doctors by City
  and Specialty" section of the source PDF).
- 1,274 deduplicated provider rows (deduplicated on NPI + Address + Specialty).

Extraction script: ..\Health_Plan_of_San_Joaquin\scrape_hpsj.py
Checkpoint/raw data: ..\Health_Plan_of_San_Joaquin\checkpoint.json
