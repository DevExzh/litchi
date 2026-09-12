#!/usr/bin/env python3
"""Compare retained replay receipts, excluding variable latency and RSS."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
baseline = ROOT / "runs/replay-20260912T222637Z.LW83lz"
replay = ROOT / "runs/root-replay"
left = [json.loads(line) for line in (baseline / "raw-receipts.jsonl").read_text().splitlines()]
right = [json.loads(line) for line in (replay / "raw-receipts.jsonl").read_text().splitlines()]
assert len(left) == len(right) == 299
ignored = {"elapsed_ns", "rss_before_bytes", "rss_after_bytes"}
for index, (before, after) in enumerate(zip(left, right)):
    assert {k: v for k, v in before.items() if k not in ignored} == {
        k: v for k, v in after.items() if k not in ignored
    }, index
for name in ("fixture-hashes.sha256", "fixture-member-hashes.sha256"):
    assert (baseline / name).read_bytes() == (replay / name).read_bytes(), name
print("PASS: 294 receipts and five correctness records; all deterministic fields and fixture/member hashes match")
