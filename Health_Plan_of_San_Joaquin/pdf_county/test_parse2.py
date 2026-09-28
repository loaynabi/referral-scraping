import sys
sys.path.insert(0, r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin")
import pdfplumber
import scrape_dsnp_and_merge as m

TESTS = [
    (r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\El_Dorado_and_Alpine_Directory_September_2026.pdf", 76, "El Dorado/Alpine", "Medicare D-SNP - El Dorado/Alpine County"),
    (r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\Stanislaus_Directory_September_2026.pdf", 80, "Stanislaus", "Medicare D-SNP - Stanislaus County"),
]
for path, pg, county, src in TESTS:
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[pg]
        rows = m.parse_page(page, pg, county, src)
        print(f"=== {county} page {pg}: {len(rows)} rows ===")
        # check word count vs row count sanity
        total_specialty_lines = 0
        for xmin, xmax in m.COL_BOUNDS:
            lines = m.get_column_lines(page, xmin, xmax)
            total_specialty_lines += sum(1 for l in lines if l.startswith('Specialty:'))
        print(f"  (specialty-line anchors found: {total_specialty_lines})")
        for r in rows:
            print(" ", r['Name'], '|', r['Credential'], '|', r['NPI'], '|', r['Office/Facility Name'], '|', r['Address'], '|', r['Office Hours'], '|', r['Languages'], '|', r['Access'], '|', r['Board Certified'])
