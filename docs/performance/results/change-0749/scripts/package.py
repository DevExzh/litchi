#!/usr/bin/env python3
"""Assemble the change 0749 evidence packet from the scratch directory.

- Raw matrix reports are kept, re-serialized without whitespace (lossless);
  `matrix/reduction-manifest.json` records every original SHA-256 and size.
- `perf stat -x,` outputs of each counter lane are bundled verbatim into one
  JSON file per lane (file name -> original text).
- Probe source, manifests, the shared lockfile and every script are copied.

Usage: package.py SCRATCH PACKET
"""

import hashlib
import json
import os
import shutil
import sys


def digest(data):
    return hashlib.sha256(data).hexdigest()


def minify_reports(source, target):
    manifest = {}
    for case in sorted(os.listdir(source)):
        directory = os.path.join(source, case)
        if not os.path.isdir(directory):
            if case.startswith("commands"):
                shutil.copy(os.path.join(source, case), os.path.join(target, case))
            continue
        os.makedirs(os.path.join(target, case), exist_ok=True)
        for name in sorted(os.listdir(directory)):
            path = os.path.join(directory, name)
            with open(path, "rb") as handle:
                raw = handle.read()
            if name.endswith(".json"):
                reduced = (json.dumps(json.loads(raw), separators=(",", ":")) + "\n").encode()
            else:
                reduced = raw
            with open(os.path.join(target, case, name), "wb") as handle:
                handle.write(reduced)
            manifest[f"{case}/{name}"] = {
                "original_sha256": digest(raw),
                "original_bytes": len(raw),
                "kept_sha256": digest(reduced),
                "kept_bytes": len(reduced),
            }
    with open(os.path.join(target, "reduction-manifest.json"), "w") as handle:
        json.dump(manifest, handle, indent=1)
        handle.write("\n")


def bundle_perf(source, target_file):
    bundle = {}
    for name in sorted(os.listdir(source)):
        if name.endswith(".perf"):
            with open(os.path.join(source, name)) as handle:
                bundle[name] = handle.read()
    with open(target_file, "w") as handle:
        json.dump(bundle, handle, indent=0)
        handle.write("\n")
    return len(bundle)


def main():
    scratch, packet = sys.argv[1], sys.argv[2]
    os.makedirs(packet, exist_ok=True)

    # Probe source, manifests and the one shared lockfile.
    probe = os.path.join(packet, "probe")
    shutil.copytree(os.path.join(scratch, "probe", "src"), os.path.join(probe, "src"), dirs_exist_ok=True)
    for arm in ("before", "after"):
        os.makedirs(os.path.join(probe, "manifests", arm), exist_ok=True)
        shutil.copy(os.path.join(scratch, "probe", "manifests", arm, "Cargo.toml"),
                    os.path.join(probe, "manifests", arm, "Cargo.toml"))
    shutil.copy(os.path.join(scratch, "probe", "manifests", "before", "Cargo.lock"),
                os.path.join(probe, "manifests", "Cargo.lock"))

    shutil.copytree(os.path.join(scratch, "scripts"), os.path.join(packet, "scripts"), dirs_exist_ok=True)
    shutil.copy(os.path.join(scratch, "build-harness.sh"), os.path.join(packet, "scripts", "build-harness.sh"))

    os.makedirs(os.path.join(packet, "matrix"), exist_ok=True)
    minify_reports(os.path.join(scratch, "matrix"), os.path.join(packet, "matrix"))
    shutil.copy(os.path.join(scratch, "analysis.json"), os.path.join(packet, "matrix", "analysis.json"))

    # The timing check under fixed glibc malloc thresholds.
    os.makedirs(os.path.join(packet, "tuned"), exist_ok=True)
    minify_reports(os.path.join(scratch, "matrix-tuned"), os.path.join(packet, "tuned"))
    shutil.copy(os.path.join(scratch, "analysis-tuned.json"), os.path.join(packet, "tuned", "analysis.json"))

    counters = os.path.join(packet, "counters")
    os.makedirs(counters, exist_ok=True)
    lanes = {
        "counters-raw": ("counters.json", "per-owner"),
        "counters-read-raw": ("counters-read.json", "read-controls"),
        "pagefaults-raw": ("pagefaults.json", "page-fault-layouts"),
        "pagefaults-tuned-raw": ("pagefaults-tuned.json", "page-fault-layouts-glibc-tuned"),
        "counters-v1-raw": ("counters-v1-digest-diluted.json", "per-owner-v1-digest-diluted"),
    }
    for raw, (summary, label) in lanes.items():
        count = bundle_perf(os.path.join(scratch, raw), os.path.join(counters, f"{label}-perf-stat.json"))
        shutil.copy(os.path.join(scratch, summary), os.path.join(counters, f"{label}.json"))
        print(f"{label}: {count} perf stat outputs")

    allocation = os.path.join(packet, "allocation")
    shutil.copytree(os.path.join(scratch, "allocation-raw"), os.path.join(allocation, "raw"), dirs_exist_ok=True)
    shutil.copy(os.path.join(scratch, "allocation.json"), os.path.join(allocation, "allocation.json"))

    shutil.copy(os.path.join(scratch, "binaries.json"), os.path.join(packet, "binaries.json"))


if __name__ == "__main__":
    main()
