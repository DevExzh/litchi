#!/usr/bin/env python3
"""Assemble the change 0749 review-round evidence under PACKET/round2.

Raw reports follow the repository's `.json.zst` convention: every timing
report is kept whole, zstd-compressed, with `reduction-manifest.json`
recording each original's SHA-256 and size. The raw `perf stat` outputs and
probe reports of the counter, corpus and allocation lanes are bundled per
lane (file name -> original text) and compressed the same way.

Usage: package2.py SCRATCH PACKET
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys


def digest(data):
    return hashlib.sha256(data).hexdigest()


def zstd(data):
    return subprocess.run(["zstd", "-19", "-q", "-c"], input=data, capture_output=True, check=True).stdout


def compress_reports(source, target):
    manifest = {}
    for case in sorted(os.listdir(source)):
        directory = os.path.join(source, case)
        if not os.path.isdir(directory):
            if case.startswith("commands"):
                shutil.copy(os.path.join(source, case), os.path.join(target, case))
            continue
        os.makedirs(os.path.join(target, case), exist_ok=True)
        for name in sorted(os.listdir(directory)):
            with open(os.path.join(directory, name), "rb") as handle:
                raw = handle.read()
            packed = zstd(raw)
            with open(os.path.join(target, case, f"{name}.zst"), "wb") as handle:
                handle.write(packed)
            manifest[f"{case}/{name}"] = {"sha256": digest(raw), "bytes": len(raw)}
    with open(os.path.join(target, "reduction-manifest.json"), "w") as handle:
        json.dump(manifest, handle, indent=1)
        handle.write("\n")


def bundle(source, target_file, suffixes):
    files = {}
    for name in sorted(os.listdir(source)):
        if name.endswith(suffixes):
            with open(os.path.join(source, name), encoding="utf-8", errors="replace") as handle:
                files[name] = handle.read()
    data = json.dumps(files, indent=0).encode()
    with open(target_file, "wb") as handle:
        handle.write(zstd(data))
    return len(files)


def main():
    scratch, packet = sys.argv[1], sys.argv[2]
    root = os.path.join(packet, "round2")
    os.makedirs(root, exist_ok=True)

    probe = os.path.join(root, "probe")
    shutil.copytree(os.path.join(scratch, "probe", "src"), os.path.join(probe, "src"), dirs_exist_ok=True)
    for arm in ("A", "B", "C"):
        os.makedirs(os.path.join(probe, "manifests", arm), exist_ok=True)
        shutil.copy(os.path.join(scratch, "probe", "manifests", arm, "Cargo.toml"),
                    os.path.join(probe, "manifests", arm, "Cargo.toml"))
    shutil.copy(os.path.join(scratch, "probe", "manifests", "A", "Cargo.lock"),
                os.path.join(probe, "manifests", "Cargo.lock"))

    scripts = os.path.join(root, "scripts")
    os.makedirs(scripts, exist_ok=True)
    for name in ("run2.py", "analyze2.py", "counters2.py", "alloc2.py", "corpus_counters.py",
                 "callgrind_scaling.py", "summarize2.py", "package2.py"):
        shutil.copy(os.path.join(scratch, "scripts", name), os.path.join(scripts, name))
    for name in ("build-probe.sh", "build-harness.sh"):
        shutil.copy(os.path.join(scratch, name), os.path.join(scripts, name))

    matrix = os.path.join(root, "matrix")
    os.makedirs(matrix, exist_ok=True)
    compress_reports(os.path.join(scratch, "matrix2"), matrix)
    shutil.copy(os.path.join(scratch, "analysis2.json"), os.path.join(matrix, "analysis.json"))
    shutil.copy(os.path.join(scratch, "analysis2-superseded-59d853618f.json"),
                os.path.join(matrix, "analysis-superseded-59d853618f.json"))

    counters = os.path.join(root, "counters")
    os.makedirs(counters, exist_ok=True)
    for raw, summary, label in (("counters2-probe-raw", "counters2-probe.json", "counters-probe"),
                                ("counters2-harness-raw", "counters2-harness.json", "counters-harness"),
                                ("counters2-tuned-raw", "counters2-tuned.json", "counters-glibc-tuned"),
                                ("quick", None, "reader-check")):
        count = bundle(os.path.join(scratch, raw), os.path.join(counters, f"{label}-perf-stat.json.zst"), (".perf",))
        if summary:
            shutil.copy(os.path.join(scratch, summary), os.path.join(counters, f"{label}.json"))
        print(f"{label}: {count} perf stat outputs")

    scaling = os.path.join(root, "scaling")
    os.makedirs(scaling, exist_ok=True)
    shutil.copy(os.path.join(scratch, "callgrind-scaling.json"), os.path.join(scaling, "callgrind-scaling.json"))
    for name in sorted(os.listdir(os.path.join(scratch, "callgrind-scaling"))):
        if name.endswith(".inclusive.txt"):
            with open(os.path.join(scratch, "callgrind-scaling", name)) as handle:
                head = "".join(handle.readlines()[:70])
            with open(os.path.join(scaling, name), "w") as handle:
                handle.write(head)

    corpus = os.path.join(root, "corpus")
    os.makedirs(corpus, exist_ok=True)
    shutil.copy(os.path.join(scratch, "corpus.json"), os.path.join(corpus, "corpus-instructions.json"))
    count = bundle(os.path.join(scratch, "corpus-raw"), os.path.join(corpus, "corpus-raw.json.zst"),
                   (".perf", ".json"))
    print(f"corpus raw: {count} files")

    allocation = os.path.join(root, "allocation")
    os.makedirs(allocation, exist_ok=True)
    shutil.copy(os.path.join(scratch, "allocation2.json"), os.path.join(allocation, "allocation.json"))
    count = bundle(os.path.join(scratch, "allocation2-raw"), os.path.join(allocation, "allocation-raw.json.zst"), (".json",))
    print(f"allocation raw: {count} reports")

    shutil.copy(os.path.join(scratch, "binaries-r2.json"), os.path.join(root, "binaries.json"))
    shutil.copy(os.path.join(scratch, "gates2", "gates.txt"), os.path.join(root, "gates.txt"))


if __name__ == "__main__":
    main()
