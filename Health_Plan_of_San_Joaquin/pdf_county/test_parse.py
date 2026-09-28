import sys
sys.path.insert(0, r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin")
import pdfplumber
import scrape_dsnp_and_merge as m

PDF = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\San_Joaquin_Directory_September_2026.pdf"
with pdfplumber.open(PDF) as pdf:
    page = pdf.pages[80]
    rows = m.parse_page(page, 80, "San Joaquin", "Medicare D-SNP - San Joaquin County")
    print("num rows:", len(rows))
    for r in rows:
        print(r)
        print()
