#!/usr/bin/env python3
"""Isolation-pair analysis: per-op inclusive/self Ir deltas and call counts.

usage: analyze.py <stem> [small=10] [large=110]
reads <cg>/<stem>-s<small>.out and <cg>/<stem>-s<large>.out
"""
import re, subprocess, sys, collections, os

CG = os.environ.get("CGDIR")
stem = sys.argv[1]
small = int(sys.argv[2]) if len(sys.argv) > 2 else 10
large = int(sys.argv[3]) if len(sys.argv) > 3 else 110
n = large - small

def annotate(path, inclusive):
    args = ["callgrind_annotate", "--threshold=100", path]
    if inclusive:
        args.insert(1, "--inclusive=yes")
    out = subprocess.run(args, capture_output=True, text=True).stdout
    total = None
    rows = {}
    for line in out.splitlines():
        m = re.match(r"^\s*([\d,]+)\s+\([\d. ]+%\)\s+(\S.*)$", line)
        if not m:
            continue
        ir = int(m.group(1).replace(",", ""))
        name = m.group(2)
        if "PROGRAM TOTALS" in name:
            total = ir
            continue
        name = re.sub(r"\s+\[.*\]$", "", name)
        name = name.split(":", 1)[-1] if name.startswith("???:") or name.startswith("./") or name.startswith("/") else name
        rows[name] = rows.get(name, 0) + ir
    return total, rows

def calls(path):
    """Aggregate call counts per callee across the whole profile."""
    names = {}
    counts = collections.Counter()
    cur_callee = None
    with open(path) as f:
        for line in f:
            if line.startswith("cfn="):
                m = re.match(r"cfn=\((\d+)\)(?: (.*))?", line.rstrip("\n"))
                if m:
                    if m.group(2):
                        names[m.group(1)] = m.group(2)
                    cur_callee = names.get(m.group(1), m.group(1))
            elif line.startswith("fn="):
                m = re.match(r"fn=\((\d+)\)(?: (.*))?", line.rstrip("\n"))
                if m and m.group(2):
                    names[m.group(1)] = m.group(2)
            elif line.startswith("calls=") and cur_callee is not None:
                counts[cur_callee] += int(line.split("=")[1].split()[0])
    return counts

ps, pl = f"{CG}/{stem}-s{small}.out", f"{CG}/{stem}-s{large}.out"
ts, incs = annotate(ps, True)
tl, incl = annotate(pl, True)
_, selfs = annotate(ps, False)
_, selfl = annotate(pl, False)
per_op = (tl - ts) / n
print(f"== {stem}: total s{small}={ts:,} s{large}={tl:,} per-op Ir={per_op:,.0f}")
print("-- inclusive per-op deltas (top 40)")
deltas = {k: (incl.get(k, 0) - incs.get(k, 0)) / n for k in set(incl) | set(incs)}
for k, v in sorted(deltas.items(), key=lambda kv: -kv[1])[:40]:
    if v <= 0:
        break
    print(f"{v:14,.0f} {100*v/per_op:6.2f}%  {k[:130]}")
print("-- self per-op deltas (top 25)")
sdeltas = {k: (selfl.get(k, 0) - selfs.get(k, 0)) / n for k in set(selfl) | set(selfs)}
for k, v in sorted(sdeltas.items(), key=lambda kv: -kv[1])[:25]:
    if v <= 0:
        break
    print(f"{v:14,.0f} {100*v/per_op:6.2f}%  {k[:130]}")
print("-- call counts per op (selected)")
cs, cl = calls(ps), calls(pl)
pat = re.compile(r"parse_impl|parse_sprms|memcpy|memset|malloc$|fc_range_to_cp_ranges|from_sprm|cascade|resolve_style_baseline|paragraph_style_on_baseline|expand_data|open_stream|read_stream|sha2|Sha256|compress|piece_for_cp|extract_from|from_utf16|PapxFkp::parse|ChpxFkp::parse|parse_container|encode_utf16|windows_1252|SlideDirectory|inspect_live|extract_all_text|extract_text|HashSet|try_reserve|finish_grow|OleFile.*open|SharedOleFile.*open|load_directory|validate_stream_allocations|collect_sector_chain|StyleSheet|parse_fkp|ParagraphProperties.*clone|from_box_in|concat")
for k in sorted(set(cs) | set(cl), key=lambda k: -(cl.get(k, 0) - cs.get(k, 0))):
    d = (cl.get(k, 0) - cs.get(k, 0)) / n
    if d >= 1 and pat.search(k):
        print(f"{d:12,.1f}  {k[:120]}")
