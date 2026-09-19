#!/usr/bin/env python3
"""Record the measured candidate and restore only its production codec."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
CODEC = "crates/litchi-ooxml-common/src/mce/codec.rs"
TESTS = "crates/litchi-ooxml-common/src/mce/tests.rs"
EXPECTED_SOURCE_FILES = 602


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


base = json.loads((P / "baseline.json").read_text())
candidate_build = next(
    row for row in json.loads((P / "builds-candidate.json").read_text())
    if row.get("label") == "native"
)
candidate = candidate_build["source_sha256"]
if len(candidate) != EXPECTED_SOURCE_FILES:
    raise AssertionError(f"candidate source census is not {EXPECTED_SOURCE_FILES} files")
if len(base["source_sha256"]) != EXPECTED_SOURCE_FILES:
    raise AssertionError(f"baseline source census is not {EXPECTED_SOURCE_FILES} files")
for name, digest in candidate.items():
    if sha(ROOT / name) != digest:
        raise AssertionError(f"candidate source changed before rejection: {name}")
if sha(P / "baseline-codec.rs.txt") != base["source_sha256"][CODEC]:
    raise AssertionError("baseline codec witness hash mismatch")
if sha(P / "baseline-tests.rs.txt") != base["source_sha256"][TESTS]:
    raise AssertionError("baseline tests witness hash mismatch")
source_diff = json.loads((P / "source-diff.json").read_text())
patch = P / "source-diff.patch"
if source_diff.get("patch") != "source-diff.patch" or sha(patch) != source_diff.get("patch_sha256"):
    raise AssertionError("source-diff receipt/hash mismatch")
if source_diff.get("baseline_head") != base["baseline_head"]:
    raise AssertionError("source-diff baseline head mismatch")
if source_diff.get("paths") != [CODEC, TESTS]:
    raise AssertionError("source-diff path census mismatch")
if (P / "rejection.json").exists() or (P / "candidate-codec.rs.txt").exists():
    raise AssertionError("rejection transition already recorded")

(P / "candidate-codec.rs.txt").write_bytes((ROOT / CODEC).read_bytes())
(ROOT / CODEC).write_bytes((P / "baseline-codec.rs.txt").read_bytes())
final = {name: sha(ROOT / name) for name in candidate}
changed_candidate = {name for name in candidate if candidate[name] != base["source_sha256"][name]}
changed_final = {name for name in final if final[name] != base["source_sha256"][name]}
if changed_candidate != {CODEC, TESTS}:
    raise AssertionError(f"candidate source transition is not codec plus tests: {sorted(changed_candidate)}")
if changed_final != {TESTS}:
    raise AssertionError(f"final source transition is not tests-only: {sorted(changed_final)}")
if final[CODEC] != base["source_sha256"][CODEC] or final[TESTS] != candidate[TESTS]:
    raise AssertionError("restored source hashes are not baseline codec plus candidate tests")

receipt = {
    "disposition": "rejected",
    "baseline_head": base["baseline_head"],
    "reason": (
        "Repeated real workflows are 1.5-3.4% slower and the tiny marked control "
        "adds about 45% (about 210 ns) plus 216 bytes of stack reservation; "
        "those costs outweigh the marker-free gains."
    ),
    "codec_path": CODEC,
    "tests_path": TESTS,
    "candidate_witness": "candidate-codec.rs.txt",
    "baseline_codec_sha256": base["source_sha256"][CODEC],
    "candidate_codec_sha256": candidate[CODEC],
    "final_codec_sha256": final[CODEC],
    "final_tests_sha256": final[TESTS],
    "source_diff_sha256": source_diff["patch_sha256"],
    "candidate_source_sha256": candidate,
    "final_source_sha256": final,
    "retained_source_changes": [TESTS],
}
(P / "rejection.json").write_text(json.dumps(receipt, indent=2) + "\n")
print("Rejected candidate preserved; production codec restored; tests retained.")
