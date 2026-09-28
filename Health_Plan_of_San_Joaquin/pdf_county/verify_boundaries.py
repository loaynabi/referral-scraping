import pdfplumber, json

boundaries = json.load(open(r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\dsnp_boundaries.json"))

FILES = {
    "El_Dorado_Alpine": r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\El_Dorado_and_Alpine_Directory_September_2026.pdf",
    "San_Joaquin": r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\San_Joaquin_Directory_September_2026.pdf",
    "Stanislaus": r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\Stanislaus_Directory_September_2026.pdf",
}

for name, path in FILES.items():
    b = boundaries[name]
    s, e = b["pcp_start"], b["pcp_end"]
    with pdfplumber.open(path) as pdf:
        for idx in [s, s+1, e-2, e-1, e]:
            if idx < 0 or idx >= len(pdf.pages):
                continue
            txt = pdf.pages[idx].extract_text() or ""
            print(f"--- {name} page idx {idx} (first 150 chars) ---")
            print(txt[:150].replace(chr(10), " | "))
            print(f"   has NPI: {'NPI:' in txt}, num_words: {len(pdf.pages[idx].extract_words())}")
    print()
