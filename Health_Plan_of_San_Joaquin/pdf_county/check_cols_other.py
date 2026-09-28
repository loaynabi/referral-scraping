import pdfplumber
from collections import Counter

FILES = {
    "El_Dorado_Alpine": (r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\El_Dorado_and_Alpine_Directory_September_2026.pdf", 75),
    "Stanislaus": (r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\Stanislaus_Directory_September_2026.pdf", 75),
}
for name, (path, pg) in FILES.items():
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[pg]
        print(name, "width:", page.width)
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        buckets = Counter(round(w['x0']/20)*20 for w in words)
        for b in sorted(buckets):
            print(" ", b, buckets[b])
