#!/usr/bin/env python3
"""Recheck retained receipts against the two frozen source checkouts."""

import argparse
import hashlib
import json
import subprocess
from pathlib import Path

import compare_manifests
import verify


def digest(data):
    return hashlib.sha256(data).hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--before-root", type=Path, required=True)
    parser.add_argument("--after-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    here = Path(__file__).resolve().parent
    snapshots = {
        "before": (args.before_root, "8b5838c591775990747b2cbce82fb2eea372b58b"),
        "after": (args.after_root, "fc44c4e6c945ab07ded7447f40670d898839eeb3"),
    }
    results = {}
    corpus = None
    normalized = None
    for side, (root, commit) in snapshots.items():
        retained = here / "results" / side
        before = retained / "source-manifest-before.txt"
        assert before.read_bytes() == (retained / "source-manifest-after.txt").read_bytes()
        count = verify.verify_manifest(before, root)
        for line in before.read_text().splitlines():
            if line.startswith("extra=\t"):
                _, shown, expected = line.split("\t")
                marker = "docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/"
                if shown.startswith(marker):
                    assert digest((here / shown.removeprefix(marker)).read_bytes()) == expected
        metadata = digest((retained / "metadata-before.json").read_bytes())
        assert f"metadata_sha256={metadata}" in before.read_text().splitlines()
        assert (retained / "metadata-before.json").read_bytes() == (retained / "metadata-after.json").read_bytes()
        label = side + "-" + commit[:9]
        verify.verify_provenance(retained, label, commit)
        verify.verify_lanes(retained, label)
        verify.verify_corpus_hashes(retained)
        verify.verify_report(retained / "report.md", verify.rows(retained))
        provenance = json.loads((retained / "provenance-verification.json").read_text())
        binary_hash = (retained / "binary.sha256").read_text().split()[0]
        assert provenance["binary_sha256"] == binary_hash
        for path in ("crates/litchi-xlsx/src/drawing/source.rs",
                     "crates/litchi-xlsx/src/drawing/source_tests.rs",
                     "crates/litchi-drawingml/src/svg_blip.rs"):
            committed = subprocess.check_output(["git", "-C", str(root), "show", f"{commit}:{path}"])
            assert digest(committed) == digest((root / path).read_bytes())
            assert digest(committed) == provenance["source_sha256"][Path(path).name]
        current_corpus = (retained / "corpus-sha256.tsv").read_bytes()
        current_normalized = compare_manifests.normalize(before)
        if corpus is not None:
            assert corpus == current_corpus
            assert normalized == current_normalized
        corpus, normalized = current_corpus, current_normalized
        receipts = [path for lane in verify.LANES for path in retained.glob(f"{lane}-p*.json")]
        samples = sum(len(json.loads(path.read_text())["samples"]) for path in receipts)
        assert len(receipts) == 117 and samples == 2340
        results[side] = {"commit": commit, "manifest_inputs_rehashed": count,
                         "manifest_sha256": digest(before.read_bytes()),
                         "binary_sha256": binary_hash, "receipts": len(receipts),
                         "samples": samples}
    record = {
        "passed": True, "snapshots": results,
        "retained_harness_extras_match_build_manifests": True,
        "corpus_sha256": digest(corpus),
        "normalized_manifest_sha256": digest("\n".join(normalized).encode()),
        "scope": "Frozen source scanner only; no lifecycle, native acceptance, or speedup claim.",
        "limitations": [
            "Binaries were hashed by the runner before target cleanup; root checks those retained hash receipts, not a second live binary.",
            "Small-output-cap lanes count any returned error; they do not assert a specific error variant.",
            "The verifier checks shared context identity and counts, but not exact namespace binding counts.",
        ],
    }
    args.output.write_text(json.dumps(record, indent=2) + "\n")
    print(json.dumps(record, indent=2))


if __name__ == "__main__":
    main()
