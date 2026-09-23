#!/usr/bin/env python3
"""Attribute a frame-pointer `perf script` capture of the owned media-rich
cross-copy lifecycle by phase, callee path and leaf work (change 0742).

Input is `perf script -F comm,tid,time,period,event,ip,sym` output of a
`perf record -e cycles --call-graph fp` run of the frame-pointer build of the
after binary. Only samples whose stack contains the timed lifecycle runner
(`run_pptx_cross_copy_lifecycle`) and not the untimed corpus construction
(`build_pptx_cross_copy_corpus`) are attributed. A sample is in the commit
phase when its stack contains `cross_copy_plan::apply_plan`, in the plan phase
when it contains `plan_cross_slide_copy`, and in publication when it contains
`OpcPackage::to_stream` directly under the runner. Within a phase, each
sample is keyed by the chain of litchi/soapberry frames below the phase root
(outermost first, truncated to DEPTH) and by a leaf category. Shares are of
sample period, not sample count.

Usage: attribute.py SCRIPT_TXT [--depth N] [--json OUT]
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
OWNERS = ("litchi_pptx::", "litchi_opc::", "soapberry_zip::", "<litchi_pptx", "<litchi_opc", "<soapberry_zip")


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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("script", type=Path)
    parser.add_argument("--depth", type=int, default=3)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    phase_period: collections.Counter = collections.Counter()
    by_path: dict = collections.defaultdict(collections.Counter)
    by_leaf: dict = collections.defaultdict(collections.Counter)
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
        root_index = None
        if any("cross_copy_plan::apply_plan" in frame for frame in frames):
            phase = "commit"
            root_index = next(i for i, f in enumerate(frames) if "cross_copy_plan::apply_plan" in f)
        elif any("plan_cross_slide_copy" in frame for frame in frames):
            phase = "plan"
            root_index = max(i for i, f in enumerate(frames) if "plan_cross_slide_copy" in f)
        elif any("OpcPackage::to_stream" in frame for frame in frames):
            phase = "publication"
            root_index = next(i for i, f in enumerate(frames) if "OpcPackage::to_stream" in f)
        else:
            phase = "other-lifecycle"
        phase_period[phase] += period
        if root_index is not None:
            chain = [short(f) for f in reversed(frames[:root_index]) if owned(f)]
            key = " > ".join(chain[: args.depth]) or "(root)"
            by_path[phase][key] += period
        by_leaf[phase][leaf_category(frames)] += period
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
        }
    text = json.dumps(report, indent=2)
    if args.json:
        args.json.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
