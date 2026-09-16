#!/usr/bin/env python3
"""Difference the probe build's per-fingerprint memo accounting at --samples 1
and --samples 3 into an exact per-lifecycle table for change 0655."""
import collections, pathlib, sys

CASES = [
    "pptx_eager_batch_edit_save",
    "pptx_eager_multi_slide_batch_edit_save",
    "pptx_slide_move_boundary_save",
    "pptx_slide_remove_boundary_save",
    "pptx_real_file_ordinary_save_edit",
    "pptx_real_file_ordinary_save_lifecycle",
]

def main(directory):
    root = pathlib.Path(directory)
    for case in CASES:
        c1 = collections.Counter((root / f"memo-per-fingerprint-{case}-s1.txt").read_text().splitlines())
        c3 = collections.Counter((root / f"memo-per-fingerprint-{case}-s3.txt").read_text().splitlines())
        delta = {k: (c3[k] - c1[k]) // 2 for k in set(c1) | set(c3) if c3[k] != c1[k]}
        print(f"## {case}: {sum(delta.values())} fingerprints per lifecycle")
        for line, count in sorted(delta.items(), key=lambda kv: -kv[1]):
            print(f"   x{count}  {line.replace('0655-fp ', '')}")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
