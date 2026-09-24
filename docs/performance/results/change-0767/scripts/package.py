#!/usr/bin/env python3
"""Change 0767: assemble the evidence packet from the scratch lanes.

Usage: package.py SCRATCH_DIR PACKET_DIR

Each lane's raw outputs are bundled as one zstd-compressed tarball with a
manifest of every member's SHA-256 and size; summaries are copied as JSON.
The verdict lane's before and after outputs are byte-identical, so one copy
is bundled and the manifest records both arms' digests. Raw Callgrind
outputs, perf.data, binaries and corpora are never copied.
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys


def digest(path):
    with open(path, "rb") as handle:
        return hashlib.sha256(handle.read()).hexdigest()


def bundle(source_root, members, archive):
    """tar + zstd of `members` (paths relative to source_root)."""
    os.makedirs(os.path.dirname(archive), exist_ok=True)
    listing = archive + ".list"
    with open(listing, "w") as handle:
        handle.write("\n".join(members) + "\n")
    subprocess.run(
        ["tar", "--sort=name", "--mtime=@0", "--owner=0", "--group=0", "--numeric-owner",
         "-C", source_root, "-T", listing, "-I", "zstd -19 -T1 -q", "-cf", archive],
        check=True)
    os.remove(listing)
    return [{"path": member, "sha256": digest(os.path.join(source_root, member)),
             "bytes": os.path.getsize(os.path.join(source_root, member))} for member in members]


def files_under(root, suffixes):
    out = []
    for directory, _, names in os.walk(root):
        for name in names:
            if name.endswith(suffixes):
                out.append(os.path.relpath(os.path.join(directory, name), root))
    return sorted(out)


def write_json(path, value):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as handle:
        json.dump(value, handle, indent=1)
        handle.write("\n")


def main():
    scratch, packet = sys.argv[1], sys.argv[2]
    # Timing matrices: the retained version and the superseded first version.
    for lane, analysis in (("matrix", "matrix-analysis.json"), ("matrix-v1", "matrix-v1-analysis.json")):
        root = f"{scratch}/{lane}"
        members = files_under(root, (".json", ".stderr"))
        members = [member for member in members if not member.startswith("commands")]
        manifest = bundle(root, members, f"{packet}/{lane}/reports.tar.zst")
        write_json(f"{packet}/{lane}/manifest.json", manifest)
        shutil.copy(f"{scratch}/{analysis}", f"{packet}/{lane}/analysis.json")
        for name in os.listdir(root):
            if name.startswith("commands"):
                shutil.copy(f"{root}/{name}", f"{packet}/{lane}/{name}")
    # Counters: the perf stat outputs (harness reports are not needed).
    for lane, summary in (("counters", "counters.json"), ("counters-v1", "counters-v1.json")):
        root = f"{scratch}/{lane}"
        manifest = bundle(root, files_under(root, (".perf",)), f"{packet}/{lane}/perf-stat.tar.zst")
        write_json(f"{packet}/{lane}/manifest.json", manifest)
        shutil.copy(f"{scratch}/{summary}", f"{packet}/{lane}/counters.json")
    # Callgrind scaling: the summary and the head of each inclusive listing.
    root = f"{scratch}/scaling"
    manifest = bundle(root, files_under(root, (".inclusive.txt",)), f"{packet}/scaling/inclusive-heads.tar.zst")
    write_json(f"{packet}/scaling/manifest.json", manifest)
    shutil.copy(f"{scratch}/callgrind-scaling.json", f"{packet}/scaling/callgrind-scaling.json")
    # The base profiles that located the terms (Callgrind inclusive heads).
    root = f"{scratch}/cg"
    heads = files_under(root, (".inclusive.txt",))
    if heads:
        manifest = bundle(root, heads, f"{packet}/profiles/base-inclusive-heads.tar.zst")
        write_json(f"{packet}/profiles/manifest.json", manifest)
    # Allocation.
    root = f"{scratch}/allocation"
    manifest = bundle(root, files_under(root, (".json",)), f"{packet}/allocation/reports.tar.zst")
    write_json(f"{packet}/allocation/manifest.json", manifest)
    shutil.copy(f"{scratch}/allocation.json", f"{packet}/allocation/allocation.json")
    # Verdicts: one copy, both arms' digests.
    root = f"{scratch}/verdicts"
    before = files_under(f"{root}/before", (".jsonl",))
    after = files_under(f"{root}/after", (".jsonl",))
    identical = before == after and all(
        digest(f"{root}/before/{name}") == digest(f"{root}/after/{name}") for name in before)
    manifest = bundle(f"{root}/before", before, f"{packet}/verdicts/verdicts.tar.zst")
    for entry in manifest:
        entry["after_sha256"] = digest(f"{root}/after/{entry['path']}")
    write_json(f"{packet}/verdicts/manifest.json", manifest)
    totals = {"files": len(before), "identical_across_arms": identical,
              "cases": 0, "opened": 0, "refused": 0, "reads_ok": 0, "reads_err": 0}
    inputs = []
    for name in before:
        with open(f"{root}/before/{name}") as handle:
            for line in handle:
                record = json.loads(line)
                if record.get("summary"):
                    for key in ("cases", "opened", "refused", "reads_ok", "reads_err"):
                        totals[key] += record[key]
                    inputs.append({"file": name, "input": record["input"],
                                   "input_sha256": record["input_sha256"], "cases": record["cases"]})
    totals["inputs"] = inputs
    write_json(f"{packet}/verdicts/verdicts-summary.json", totals)


if __name__ == "__main__":
    main()
