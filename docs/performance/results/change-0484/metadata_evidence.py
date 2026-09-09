#!/usr/bin/env python3
"""Verify metadata profiles and retain raw traces through lossless gzip custody."""
import argparse
import gzip
import hashlib
from pathlib import Path
from common import ROOT, meta, now, read, write

ATTEMPT = ROOT / "input-metadata-profiles/metadata2"
RECEIPT = ATTEMPT / "compressed-custody.json"


def check(compact=False):
    result = read(ATTEMPT / "result.json")
    assert result["passed"] and result["binary_unchanged"] and not result["failures"]
    assert result["status"] == "ok" and len(result["receipts"]) == 6
    assert result["driver"]["sha256"] == meta(ROOT / "profile_input_metadata.py")["sha256"]
    entries = []
    for record in result["receipts"]:
        assert record["status"] == "ok" and record["passed"]
        binding = record["receipt"]
        path = ROOT / binding["path"]
        assert meta(path) == {k: binding[k] for k in ("bytes", "sha256")}
        receipt = read(path)
        assert receipt["passed"] and receipt["validation_status"] == "ok"
        trace = receipt.get("raw_trace_binding")
        raw_path = None
        if trace:
            raw = trace["raw_before_compression"]
            raw_path = ROOT / raw["path"]
            compressed = trace["compressed"]
            compressed_path = ROOT / compressed["path"]
            assert meta(compressed_path) == {k: compressed[k] for k in ("bytes", "sha256")}
            digest = hashlib.sha256()
            size = source_statx = all_statx = 0
            with gzip.open(compressed_path, "rb") as stream:
                for line in stream:
                    digest.update(line)
                    size += len(line)
                    if b"statx(" in line:
                        all_statx += 1
                        if b"/axis-input/s64-a16384-short-c64-source.docx>" in line:
                            source_statx += 1
            assert {"bytes": size, "sha256": digest.hexdigest()} == {k: raw[k] for k in ("bytes", "sha256")}
            entries.append({"raw": raw, "compressed": compressed,
                            "receipt": binding, "raw_statx_calls": all_statx,
                            "exact_source_descriptor_statx_calls": source_statx,
                            "lossless_verified": True, "raw_removed": True})
        for name, artifact in receipt["artifacts"].items():
            artifact_path = path.parent / name
            if artifact_path == raw_path and not artifact_path.exists():
                assert not compact and RECEIPT.is_file()
                continue
            if artifact.get("present"):
                assert meta(artifact_path) == {k: artifact[k] for k in ("bytes", "sha256")}
        if raw_path is not None and compact:
            assert raw_path.is_file() and not raw_path.is_symlink()
            raw_path.unlink()
    if compact:
        write(RECEIPT, {"schema": "docx-input-metadata-compressed-custody-v1",
                        "status": "pass", "finished_utc": now(),
                        "scope": "Whole diagnostic child including one warmup, one sample and setup; exact descriptor annotation attribution.",
                        "driver": meta(Path(__file__)), "entries": entries})
    else:
        retained = read(RECEIPT)
        assert retained["status"] == "pass" and retained["entries"] == entries
        assert retained["driver"] == meta(Path(__file__))
    return entries


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("compact", "verify"))
    args = parser.parse_args()
    if args.command == "compact":
        assert not RECEIPT.exists()
    for entry in check(args.command == "compact"):
        print(entry["raw"]["path"], entry["raw_statx_calls"],
              entry["exact_source_descriptor_statx_calls"], flush=True)
