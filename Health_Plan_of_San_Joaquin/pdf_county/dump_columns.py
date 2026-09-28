import pdfplumber

PDF = r"C:\Users\DamonHollister\Desktop\referral scraping\Health_Plan_of_San_Joaquin\pdf_county\San_Joaquin_Directory_September_2026.pdf"
COL_BOUNDS = [(0, 175), (175, 385), (385, 999)]

def get_column_lines(page, x_min, x_max):
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    col_words = [w for w in words if x_min <= w['x0'] < x_max]
    rows = {}
    for w in col_words:
        key = round(w['top'])
        rows.setdefault(key, []).append(w)
    lines = []
    for top in sorted(rows.keys()):
        ws = sorted(rows[top], key=lambda w: w['x0'])
        text = " ".join(w['text'] for w in ws)
        lines.append((top, text))
    return lines

with pdfplumber.open(PDF) as pdf:
    page = pdf.pages[80]
    for ci, (xmin, xmax) in enumerate(COL_BOUNDS):
        print(f"=== COLUMN {ci} ({xmin}-{xmax}) ===")
        lines = get_column_lines(page, xmin, xmax)
        for top, text in lines:
            print(f"  [{top}] {text}")
