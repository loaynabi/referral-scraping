import json, sys, os

CHECKPOINT = os.path.join(os.path.dirname(__file__), "checkpoint_raw.json")

def load_checkpoint():
    if os.path.exists(CHECKPOINT):
        with open(CHECKPOINT, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_checkpoint(rows):
    with open(CHECKPOINT, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)

def ingest(tool_result_path, tag=None):
    with open(tool_result_path, "r", encoding="utf-8") as f:
        envelope = json.load(f)
    text = envelope[0]["text"]
    # Strip the "(captured at origin ...)" trailer the browser tool appends
    marker = "\n\n(captured at origin"
    idx = text.rfind(marker)
    if idx != -1:
        text = text[:idx]
    # text is now a JSON-encoded string (JS did JSON.stringify, tool wrapped again) -> unescape once
    inner = json.loads(text)
    rows = json.loads(inner)
    existing = load_checkpoint()
    existing.extend(rows)
    save_checkpoint(existing)
    print(f"Ingested {len(rows)} rows from {tool_result_path}. Total checkpoint size: {len(existing)}")

if __name__ == "__main__":
    for p in sys.argv[1:]:
        ingest(p)
