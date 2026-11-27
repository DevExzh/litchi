#!/usr/bin/env python3
"""Regenerate the `officecli-*.json` seeds of `corpus/batch_payload/`.

Each seed is the first six items of the first batch JSON heredoc in one of the
OfficeCLI example scripts tracked under `test-data/officecli/examples/`
(OfficeCLI, Apache-2.0; provenance in `test-data/officecli/PROVENANCE.md`).
The hand-written seeds beside them come from the `litchi-ops` batch decoder
tests (`src/batch/tests.rs`) and are not regenerated.

Run from the repository root: python3 crates/litchi-ops/fuzz/harvest_batch_seeds.py
"""

import json
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[3]
EXAMPLES = ROOT / "test-data" / "officecli" / "examples"
OUT = pathlib.Path(__file__).resolve().parent / "corpus" / "batch_payload"
HEREDOC = re.compile(r"<< ?'?(\w+)'?[^\n]*\n(\[.*?\])\s*\n\1", re.S)

for script in sorted(EXAMPLES.rglob("*.sh")):
    for match in HEREDOC.finditer(script.read_text(encoding="utf-8")):
        try:
            items = json.loads(match.group(2))
        except json.JSONDecodeError:
            continue
        name = "officecli-" + script.parent.name.replace("--", "-") + ".json"
        text = json.dumps(items[:6], separators=(",", ":"), ensure_ascii=False)
        (OUT / name).write_text(text, encoding="utf-8")
        print(f"{name}: {len(text.encode())} bytes from {script.relative_to(ROOT)}")
        break
