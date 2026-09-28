import pdfplumber
import re, json, sys

FILES = {
    "El_Dorado_Alpine": r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\El_Dorado_and_Alpine_Directory_September_2026.pdf",
    "San_Joaquin": r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\San_Joaquin_Directory_September_2026.pdf",
    "Stanislaus": r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\Stanislaus_Directory_September_2026.pdf",
}

PCP_HDR = re.compile(r"C1\.\s*Primary Care Providers", re.I)
SPEC_HDR = re.compile(r"C2\.\s*Specialists", re.I)

results = {}
for name, path in FILES.items():
    print(f"=== {name} ===")
    with pdfplumber.open(path) as pdf:
        n = len(pdf.pages)
        print("n_pages:", n)
        pcp_start = None
        pcp_end = None
        for i in range(n):
            try:
                txt = pdf.pages[i].extract_text() or ""
            except Exception as e:
                print(f"  page {i} extract error: {e}")
                continue
            first_200 = txt[:200]
            has_pcp = bool(PCP_HDR.search(first_200))
            has_spec = bool(SPEC_HDR.search(first_200))
            if has_pcp and pcp_start is None:
                pcp_start = i
                print(f"  PCP header first seen at pdfplumber index {i}")
            if has_spec and pcp_start is not None and pcp_end is None:
                pcp_end = i
                print(f"  SPEC header first seen at pdfplumber index {i} -> pcp_end (exclusive)")
                break
        results[name] = {"n_pages": n, "pcp_start": pcp_start, "pcp_end": pcp_end}
        print(f"  RESULT: start={pcp_start} end_exclusive={pcp_end}")

out_path = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\dsnp_boundaries.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2)
print("Wrote", out_path)
