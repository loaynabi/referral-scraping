import pdfplumber

PDF = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\San_Joaquin_Directory_September_2026.pdf"
with pdfplumber.open(PDF) as pdf:
    page = pdf.pages[75]
    print("page width:", page.width)
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    xs = sorted(set(round(w['x0']) for w in words))
    print("distinct x0 (rounded), count:", len(xs))
    print(xs)
    # histogram bucket by 10s
    from collections import Counter
    buckets = Counter(round(w['x0']/20)*20 for w in words)
    for b in sorted(buckets):
        print(b, buckets[b])
