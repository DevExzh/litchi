#!/usr/bin/env python3
"""Prune S10 harness reports to the fields analyze_s10.py reads.

The full reports (environment, corpus catalogs, per-sample filesystem
evidence) stay outside the repository. This keeps, per child report:
- results[]: case, cache_state, corpus {name, archive_sha256, archive_bytes},
  elapsed_ns.samples, operation_metrics.process.peak_rss_bytes.values, and
  source.unified_ops {projection_sha256, observation, legs_agree} /
  source.facade_ole2.observation;
- filesystem_evidence[]: case, corpus.name, and per sample only the digest
  fields (`*_sha256`, `*digest`, docx_timed_full_text_bytes);
- binary_identity.

Output: OUT/pruned-<group>.json.gz, a JSON object {child stem: pruned report},
written with gzip -9 semantics and no timestamp (mtime 0).

usage: prune_reports.py RAW_DIR OUT_DIR
"""

import gzip
import json
import os
import sys
from collections import defaultdict


def prune(report):
    results = []
    for result in report["results"]:
        corpus = result["corpus"]
        entry = {
            "case": result["case"],
            "cache_state": result.get("cache_state"),
            "corpus": {key: corpus.get(key) for key in ("name", "archive_sha256", "archive_bytes")},
            "elapsed_ns": {"samples": result["elapsed_ns"]["samples"]},
        }
        process = ((result.get("operation_metrics") or {}).get("process") or {})
        peak = (process.get("peak_rss_bytes") or {}).get("values")
        if peak:
            entry["operation_metrics"] = {"process": {"peak_rss_bytes": {"values": peak}}}
        source = result.get("source") or {}
        kept = {}
        unified = source.get("unified_ops") or {}
        if unified:
            kept["unified_ops"] = {key: unified[key] for key in
                                   ("projection_sha256", "observation", "legs_agree") if key in unified}
        facade = source.get("facade_ole2") or {}
        if "observation" in facade:
            kept["facade_ole2"] = {"observation": facade["observation"]}
        if kept:
            entry["source"] = kept
        results.append(entry)
    evidence = []
    for item in report.get("filesystem_evidence") or []:
        samples = []
        for sample in item["samples"]:
            samples.append({key: value for key, value in sample.items()
                            if value is not None and (key.endswith("_sha256") or key.endswith("digest")
                                                      or key == "docx_timed_full_text_bytes")})
        evidence.append({"case": item["case"], "corpus": {"name": item["corpus"].get("name")},
                         "samples": samples})
    return {"binary_identity": report.get("binary_identity"), "results": results,
            "filesystem_evidence": evidence}


def main():
    raw, out = sys.argv[1], sys.argv[2]
    bundles = defaultdict(dict)
    for name in sorted(os.listdir(raw)):
        if (".corpus-manifest.json" in name or ".quota-failed" in name
                or not (name.endswith(".json") or name.endswith(".json.gz"))):
            continue
        stem = name[:-8] if name.endswith(".json.gz") else name[:-5]
        group = stem.split("-", 2)[2]
        opener = gzip.open if name.endswith(".gz") else open
        with opener(os.path.join(raw, name), "rt") as handle:
            bundles[group][stem] = prune(json.load(handle))
    os.makedirs(out, exist_ok=True)
    for group, bundle in sorted(bundles.items()):
        data = json.dumps(bundle, sort_keys=True, separators=(",", ":")).encode()
        path = os.path.join(out, f"pruned-{group}.json.gz")
        with open(path, "wb") as handle:
            with gzip.GzipFile(fileobj=handle, mode="wb", compresslevel=9, mtime=0, filename="") as gz:
                gz.write(data)
        print(f"{path}: {len(bundle)} children, {os.path.getsize(path)} bytes")


if __name__ == "__main__":
    main()
