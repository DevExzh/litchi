#!/usr/bin/env python3
"""Compare the base and branch probe outputs line by line and summarize which
routes the inputs reached."""
import collections, json, re, sys
summary = {}
for mode in ("docx", "snapxml", "pptx"):
    before = [json.loads(l) for l in open(f"before-{mode}.jsonl")]
    after = [json.loads(l) for l in open(f"after-{mode}.jsonl")]
    assert len(before) == len(after)
    diffs = [(b, a) for b, a in zip(before, after) if b != a]
    kinds = collections.Counter()
    for row in before:
        for key, value in row.items():
            if key == "file" or key == "edit_paragraphs":
                continue
            v = str(value)
            field = re.sub(r"_\d+$", "", key)
            if re.match(r"^(ok|len=|bytes=|n=|p=)", v) or v.startswith("changed="):
                status = "ok"
            else:
                match = re.search(r"(sb-[a-z]+-err|open-err|edit-err|stage-err|publish-err|save-err|second-[a-z]+-err|inv-err|err|open)", v)
                status = match.group(1) if match else v.split(" ")[0][:20]
                if "sb-sha:" in v:
                    status = "ok"
            kinds[(field, status)] += 1
    summary[mode] = {"rows": len(before), "differing_rows": len(diffs), "outcomes": {f"{k[0]}:{k[1]}": v for k, v in sorted(kinds.items())}}
    for b, a in diffs[:5]:
        for key in b:
            if b.get(key) != a.get(key):
                print(mode, b["file"], key, "\n  before:", str(b.get(key))[:300], "\n  after: ", str(a.get(key))[:300])
json.dump(summary, open("summary.json", "w"), indent=1, sort_keys=True)
for mode, s in summary.items():
    print(mode, "rows", s["rows"], "differing", s["differing_rows"])
    print("  ", s["outcomes"])
