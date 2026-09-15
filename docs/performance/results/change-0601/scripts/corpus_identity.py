#!/usr/bin/env python3
"""Extract the corpus identities from one harness report, for the determinism diff."""

from __future__ import annotations

import json
import sys


def main() -> int:
    document = json.load(open(sys.argv[1], encoding="utf-8"))
    seen: dict[str, dict] = {}
    for result in document["results"]:
        corpus = result["corpus"]
        seen[corpus["name"]] = {
            "archive_bytes": corpus["archive_bytes"],
            "archive_sha256": corpus["archive_sha256"],
            "archive_member_count": corpus["archive_member_count"],
            "part_bytes": corpus["uncompressed_payload_bytes"],
            "target_entry": corpus["target_entry"],
            "target_payload_sha256": corpus["target_payload_sha256"],
        }
    with open(sys.argv[2], "w", encoding="utf-8") as handle:
        json.dump(seen, handle, indent=2, sort_keys=True)
        handle.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
