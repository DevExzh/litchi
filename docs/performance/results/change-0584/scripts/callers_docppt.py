#!/usr/bin/env python3
"""Attribute __memcpy_avx_unaligned_erms / __memset_avx2_unaligned_erms to callers.

Reads `callgrind_annotate --tree=caller` output for the small and large leg of
each isolation pair, differences the caller->callee edges and divides by the
sample delta, giving Ir and call count per operation charged to each caller.
"""
import re
import pathlib

ROOT = pathlib.Path("/tmp/claude-1001/-home-zhuhe-code-litchi/14c44904-927d-4351-97ac-5611bafb5316/scratchpad/out-docppt/ann")
SMALL, LARGE = 20, 320
SPAN = LARGE - SMALL

CELLS = [
    ("docbig", "open"), ("docbig", "text"),
    ("docmid", "open"), ("docmid", "text"),
    ("docsmall", "open"), ("docsmall", "text"),
    ("pptbig", "open"), ("pptbig", "text"),
    ("pptmid", "open"), ("pptmid", "text"),
    ("pptsmall", "open"), ("pptsmall", "text"),
]

NUM = re.compile(r"^\s*([\d,]+) \(\s*[\d.]+%\)\s+(.*)$")
OBJ = re.compile(r"\s*\[[^\]]*\]\s*$")
FILE_PREFIX = re.compile(r"^.*?:(?!:)")


def clean(text):
    return FILE_PREFIX.sub("", OBJ.sub("", text).strip(), count=1)


def edges(path):
    """{(caller, callee): (calls, ir)} from a --tree=caller annotation."""
    out, pending = {}, []
    if not path.exists():
        return out
    for line in path.read_text(errors="replace").splitlines():
        match = NUM.match(line)
        if not match:
            pending = []
            continue
        ir = int(match.group(1).replace(",", ""))
        rest = match.group(2)
        if rest.startswith("*"):
            callee = clean(rest[1:])
            for caller, calls, cir in pending:
                key = (caller, callee)
                prev = out.get(key, (0, 0))
                out[key] = (prev[0] + calls, prev[1] + cir)
            pending = []
        elif rest.startswith("<"):
            body = rest[1:].strip()
            cm = re.search(r"\((\d[\d,]*)x\)\s*(\[[^\]]*\])?\s*$", body)
            calls = int(cm.group(1).replace(",", "")) if cm else 0
            body = re.sub(r"\s*\(\d[\d,]*x\)\s*(\[[^\]]*\])?\s*$", "", body)
            pending.append((clean(body), calls, ir))
        else:
            pending = []
    return out


def program_total(path):
    for line in path.read_text(errors="replace").splitlines():
        if "PROGRAM TOTALS" in line:
            return int(re.match(r"^\s*([\d,]+)", line).group(1).replace(",", ""))
    return 0


for stem, op in CELLS:
    small = ROOT / f"tree-{stem}-{op}-s{SMALL}.txt"
    large = ROOT / f"tree-{stem}-{op}-s{LARGE}.txt"
    a, b = edges(small), edges(large)
    if not b:
        print(f"\n### {stem}/{op}: MISSING tree annotation")
        continue
    whole = (program_total(ROOT / f"self-{stem}-{op}-s{LARGE}.txt")
             - program_total(ROOT / f"self-{stem}-{op}-s{SMALL}.txt")) / SPAN
    rows = []
    for key in set(a) | set(b):
        if "memcpy_avx" not in key[1] and "memset_avx2" not in key[1]:
            continue
        calls = (b.get(key, (0, 0))[0] - a.get(key, (0, 0))[0]) / SPAN
        ir = (b.get(key, (0, 0))[1] - a.get(key, (0, 0))[1]) / SPAN
        if ir > 20:
            rows.append((ir, calls, key[0], "memcpy" if "memcpy" in key[1] else "memset"))
    rows.sort(reverse=True)
    print(f"\n### {stem}/{op} -- memcpy/memset by CALLER (whole op = {whole:,.0f} Ir)")
    print(f"  {'Ir/op':>11} {'%whole':>7} {'calls/op':>9}  fn      caller")
    covered = 0.0
    for ir, calls, caller, kind in rows[:14]:
        covered += ir
        print(f"  {ir:>11,.0f} {ir/whole*100:>6.2f}% {calls:>9,.1f}  {kind:<7} {caller[:105]}")
    tail = sum(r[0] for r in rows) - covered
    print(f"  {'(other edges)':>11} {tail:,.0f} Ir/op ; attributed total {sum(r[0] for r in rows):,.0f} Ir/op"
          f" = {sum(r[0] for r in rows)/whole*100:.2f}% of the whole operation")
