#!/usr/bin/env python3
"""Focused checks for the profile interpretation and deterministic projection."""
from collections import Counter
from pathlib import Path
import json
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "profile"))
import summarize

left = Counter({("b", "elf"): 10, ("a", "elf"): 10})
right = Counter({("a", "elf"): 10, ("b", "elf"): 10})
assert summarize._counter_rows(left, 20) == summarize._counter_rows(right, 20)
assert not summarize._match_class(("shared_strings::shape", "elf"), "sha")
assert summarize._match_class(("sha2::sha256::compress", "elf"), "sha")
assert not summarize._match_class(("flate2::mem::Decompress::decompress", "elf"), "deflate")
weighted = summarize._weighted([{
    "period": 10, "event": "cycles:u",
    "frames": [{"symbol": symbol, "dso": "elf"} for symbol in
               ("zlib_rs::deflate::deflate", "DeflateEncoder::write", "main")],
}])
assert weighted["families"]["deflate"]["inclusive_percent"] == 100.0
assert weighted["families"]["deflate"]["inclusive_period"] == 10
print(json.dumps({"status": "pass", "checks": 5,
                  "scope": "stable ties, precise families, once-per-stack family weight"}))
