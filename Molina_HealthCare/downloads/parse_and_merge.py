import json, re, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOWNLOADS = os.path.join(BASE, "downloads")

def parse_sacramento(path):
    with open(path, encoding="utf-8") as f:
        text = f.read()
    blocks = [b.strip() for b in text.split("\n\n") if b.strip()]
    records = []
    for b in blocks:
        lines = [l.rstrip() for l in b.split("\n") if l.strip()]
        if len(lines) < 5:
            continue
        name = lines[0]
        specialty = lines[1]
        npi_line = lines[2]
        npi = npi_line.replace("NPI:", "").strip()
        medical_group = lines[3]
        addr_line = lines[4]
        phone_line = lines[5] if len(lines) > 5 else ""
        phone = phone_line.replace("Phone:", "").strip()
        # address line format: "street, City, ST ZIP"
        m = re.match(r"^(.*),\s*([A-Za-z .]+),\s*([A-Z]{2})\s*(\d{5})$", addr_line)
        if m:
            street, city, state, zip_ = m.groups()
        else:
            street, city, state, zip_ = addr_line, "Sacramento", "CA", ""
        # strip " PCP" suffix from name if present (site label, not part of legal name)
        clean_name = re.sub(r"\s+PCP$", "", name)
        records.append({
            "name": clean_name,
            "is_pcp": True,
            "primary_specialty": specialty,
            "npi": npi,
            "medical_group": medical_group,
            "address": street.strip(),
            "city": city.strip(),
            "state": state.strip(),
            "zip": zip_.strip(),
            "phone": phone,
            "accepting_new_patients": "",
            "source_search_zip": "95814",
            "source_search_city": "Sacramento",
        })
    return records

def dedup_key(r):
    return (r.get("npi") or "", (r.get("address") or "").lower(), (r.get("primary_specialty") or "").lower())

def main():
    checkpoint_path = os.path.join(BASE, "checkpoint.json")
    with open(checkpoint_path, encoding="utf-8") as f:
        existing = json.load(f)

    sac_new = parse_sacramento(os.path.join(DOWNLOADS, "sacramento_raw.txt"))

    merged = list(existing)
    seen = set(dedup_key(r) for r in existing)
    added = 0
    for r in sac_new:
        k = dedup_key(r)
        if k in seen:
            continue
        seen.add(k)
        merged.append(r)
        added += 1

    print(f"existing={len(existing)} sac_new_parsed={len(sac_new)} added={added} merged_total={len(merged)}")

    with open(checkpoint_path, "w", encoding="utf-8") as f:
        json.dump(merged, f, indent=2)

if __name__ == "__main__":
    main()
