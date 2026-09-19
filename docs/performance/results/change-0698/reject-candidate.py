#!/usr/bin/env python3
"""Preserve the measured candidate and restore only its production codec."""
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
CODEC = "crates/litchi-ooxml-common/src/mce/codec.rs"
TESTS = "crates/litchi-ooxml-common/src/mce/tests.rs"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


base = json.loads((P / "baseline.json").read_text())
build = next(row for row in json.loads((P / "builds-candidate.json").read_text())
             if row["label"] == "native")
candidate = build["source_sha256"]
for name, digest in candidate.items():
    assert sha(ROOT / name) == digest, name
assert sha(P / "baseline-codec.rs.txt") == base["source_sha256"][CODEC]
assert sha(P / "source-diff.patch") == json.loads((P / "source-diff.json").read_text())["patch_sha256"]
assert not (P / "rejection.json").exists(), "rejection transition already recorded"
(P / "candidate-codec.rs.txt").write_bytes((ROOT / CODEC).read_bytes())
(ROOT / CODEC).write_bytes((P / "baseline-codec.rs.txt").read_bytes())
final = {name: sha(ROOT / name) for name in candidate}
assert {name for name in final if final[name] != candidate[name]} == {CODEC}
assert {name for name in final if final[name] != base["source_sha256"][name]} == {TESTS}
receipt = {
    "disposition": "rejected",
    "baseline_head": base["baseline_head"],
    "reason": "Reachable early-name refusal exceeds the +5% review threshold in both independent measurement rounds.",
    "codec_path": CODEC,
    "tests_path": TESTS,
    "candidate_witness": "candidate-codec.rs.txt",
    "baseline_codec_sha256": base["source_sha256"][CODEC],
    "candidate_codec_sha256": candidate[CODEC],
    "final_codec_sha256": final[CODEC],
    "final_tests_sha256": final[TESTS],
    "source_diff_sha256": sha(P / "source-diff.patch"),
    "candidate_source_sha256": candidate,
    "final_source_sha256": final,
    "retained_source_changes": [TESTS],
}
(P / "rejection.json").write_text(json.dumps(receipt, indent=2) + "\n")
print("Rejected candidate preserved; production codec restored; regression tests retained.")
