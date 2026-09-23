#!/usr/bin/env python3
"""Attribute a frame-pointer `perf script` capture of the owned media-rich
cross-copy lifecycle by phase, callee path and SHA-256 call site (change 0751).

Adapted from change 0742's `attribution/attribute.py`. Input is
`perf script -F comm,tid,time,period,event,ip,sym` output of a
`perf record -e cycles --call-graph fp` run of a frame-pointer build of the
harness. Only samples whose stack contains the timed lifecycle runner
(`run_pptx_cross_copy_lifecycle`) and not the untimed corpus construction
(`build_pptx_cross_copy_corpus`) are attributed.

Phases (first match wins):
  commit       stack contains `Package::apply_cross_slide_copy_plan` (release
               builds inline `cross_copy_plan::apply_plan` into it) or
               `cross_copy_plan::apply_plan`;
  plan         stack contains `plan_cross_slide_copy`;
  publication  stack contains `OpcPackage::to_stream` directly under the runner;
  capture      stack contains `Package::opened_presentation`;
  open         stack contains `Package::from_vec`;
  other        anything else inside the runner.

Within a phase every sample is keyed by the chain of litchi/soapberry frames
below the phase root (outermost first, truncated to DEPTH) and by a leaf
category. SHA-256 leaf samples are additionally keyed by their hash site: the
chain of owned frames from the phase root down to the innermost owned frame
(truncated to SITE_DEPTH, innermost kept). Shares are of sample period, not
sample count.

Usage: attribute.py SCRIPT_TXT [--depth N] [--site-depth N] [--json OUT]
"""

from __future__ import annotations

import argparse
import collections
import json
import re
import sys
from pathlib import Path

HEADER = re.compile(r"^\S.*?\s(\d+)\s+\S+:\s*$")
LEAF_CATEGORIES = (
    ("sha256", ("sha2::",)),
    ("memcpy", ("memmove", "memcpy")),
    ("memcmp", ("memcmp", "bcmp")),
    ("memset", ("memset",)),
    ("crc32", ("crc32",)),
    ("inflate", ("zlib_rs::inflate", "inflate")),
    ("deflate", ("zlib_rs::deflate",)),
    ("allocator", ("malloc", "_int_free", "free", "realloc", "calloc", "munmap", "mmap")),
)
OWNERS = (
    "litchi_pptx::",
    "litchi_opc::",
    "soapberry_zip::",
    "<litchi_pptx",
    "<litchi_opc",
    "<soapberry_zip",
)
PHASES = (
    ("commit", ("Package::apply_cross_slide_copy_plan", "cross_copy_plan::apply_plan")),
    ("plan", ("plan_cross_slide_copy",)),
)


def symbol(line: str) -> str:
    parts = line.strip().split(maxsplit=1)
    return parts[1] if len(parts) == 2 else "[unknown]"


def leaf_category(frames: list[str]) -> str:
    leaf = frames[0]
    if leaf.startswith("ffffffff") or leaf == "[unknown]":
        return "kernel-or-unknown"
    for category, needles in LEAF_CATEGORIES:
        if any(needle in leaf for needle in needles):
            return category
    return "other"


def owned(frame: str) -> bool:
    return any(frame.startswith(owner) for owner in OWNERS)


def short(frame: str) -> str:
    frame = re.sub(r"::h[0-9a-f]{16}$", "", frame)
    frame = re.sub(r"<impl [^>]*>::", "", frame)
    frame = frame.replace("::{{closure}}", "{closure}")
    return frame.split("::")[-1] if "::" in frame else frame


def parse(path: Path):
    block: list[str] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if line.strip():
            block.append(line)
            continue
        if block:
            yield block
            block = []
    if block:
        yield block


def classify(frames: list[str]) -> tuple[str, int | None]:
    for phase, needles in PHASES:
        hits = [i for i, f in enumerate(frames) if any(n in f for n in needles)]
        if hits:
            # Outermost matching frame is the phase root.
            return phase, max(hits)
    runner = next(i for i, f in enumerate(frames) if "run_pptx_cross_copy_lifecycle" in f)
    below = frames[:runner]
    if below and any("OpcPackage::to_stream" in f for f in below[-3:]):
        return "publication", max(i for i, f in enumerate(below) if "OpcPackage::to_stream" in f)
    for phase, needle in (
        ("capture", "Package::opened_presentation"),
        ("open", "Package::from_vec"),
    ):
        hits = [i for i, f in enumerate(below) if needle in f]
        if hits:
            return phase, max(hits)
    return "other", None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("script", type=Path)
    parser.add_argument("--depth", type=int, default=3)
    parser.add_argument("--site-depth", type=int, default=4)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    phase_period: collections.Counter = collections.Counter()
    by_path: dict = collections.defaultdict(collections.Counter)
    by_leaf: dict = collections.defaultdict(collections.Counter)
    by_site: dict = collections.defaultdict(collections.Counter)
    total = 0
    for block in parse(args.script):
        match = HEADER.match(block[0])
        if not match:
            continue
        period = int(match.group(1))
        frames = [symbol(line) for line in block[1:]]
        if not frames:
            continue
        joined = "\n".join(frames)
        if "run_pptx_cross_copy_lifecycle" not in joined or "build_pptx_cross_copy_corpus" in joined:
            continue
        total += period
        phase, root_index = classify(frames)
        phase_period[phase] += period
        leaf = leaf_category(frames)
        by_leaf[phase][leaf] += period
        if root_index is not None:
            chain = [short(f) for f in reversed(frames[:root_index]) if owned(f)]
            key = " > ".join(chain[: args.depth]) or "(root)"
            by_path[phase][key] += period
            if leaf == "sha256":
                site = chain[-args.site_depth :] if chain else ["(root)"]
                by_site[phase][" > ".join(site)] += period
    report: dict = {"lifecycle_period": total, "phases": {}}
    for phase, period in phase_period.most_common():
        report["phases"][phase] = {
            "share_of_lifecycle": round(period / total, 4),
            "leaf_shares": {
                leaf: round(value / period, 4) for leaf, value in by_leaf[phase].most_common()
            },
            "path_shares": {
                path: round(value / period, 4)
                for path, value in by_path[phase].most_common(25)
            },
            "sha256_site_shares_of_phase": {
                site: round(value / period, 4)
                for site, value in by_site[phase].most_common(25)
            },
        }
    text = json.dumps(report, indent=2)
    if args.json:
        args.json.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
