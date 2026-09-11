#!/usr/bin/env python3
"""Verify a copied receipt bundle without requiring its frozen source checkout."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from verify import verify_corpus_hashes, verify_lanes, verify_provenance, verify_report
from summarize import rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--expected-commit", required=True)
    args = parser.parse_args()
    results = args.results.resolve()
    manifest_before = results / "source-manifest-before.txt"
    manifest_after = results / "source-manifest-after.txt"
    if manifest_before.read_bytes() != manifest_after.read_bytes():
        raise SystemExit("source manifest changed during profile")
    verify_provenance(results, args.snapshot, args.expected_commit)
    verify_lanes(results, args.snapshot)
    verify_corpus_hashes(results)
    verify_report(results / "report.md", rows(results))
    print(
        json.dumps(
            {
                "passed": True,
                "results": str(results),
                "lanes": 39,
                "processes_per_lane": 3,
                "samples_per_process": 20,
                "source_manifest_byte_stable": True,
                "source_manifest_rehash": "use the detached worktree verify.py for package-file hash revalidation",
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
