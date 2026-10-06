#!/usr/bin/env python3
"""Regenerate the compact JSONL source listing for 0931 custody."""
import argparse, hashlib, json
from pathlib import Path

def digest(path):
    h=hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024*1024), b""):
            h.update(chunk)
    return h.hexdigest()

parser=argparse.ArgumentParser()
parser.add_argument("--root", type=Path, required=True)
parser.add_argument("--output", type=Path, required=True)
args=parser.parse_args()
root=args.root.resolve(); output=args.output.resolve()
rows=[]
for path in sorted(root.rglob("*")):
    if path.is_file() and not path.is_symlink():
        rows.append({"bytes":path.stat().st_size,"path":path.relative_to(root).as_posix(),"sha256":digest(path)})
with output.open("x", encoding="utf-8") as stream:
    for row in rows:
        stream.write(json.dumps(row, sort_keys=True, separators=(",", ":"))+"\n")
