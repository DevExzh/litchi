import pathlib, re, collections, sys
p = pathlib.Path(sys.argv[1])
phase="pre"; cur=None; out=collections.defaultdict(collections.Counter)
for line in p.read_text(errors="replace").splitlines():
    if "#MARK#" in line:
        phase=line.split("#")[2]; continue
    if "===== read_at" in line: cur=[]; continue
    if "=====END=====" in line and cur is not None:
        out[phase][" <- ".join(cur[:6]) or "(none)"] += 1; cur=None; continue
    if cur is None: continue
    m = re.match(r"^#\d+\s+(?:0x\S+ in )?(\S+)\s*\(.*\) at (.*?)(?::(\d+))?$", line)
    if m:
        fn = m.group(1)
        if (fn.startswith(("soapberry_","litchi_")) and "read_at" not in fn
                and "validate_range" not in fn and "read_source_at" not in fn):
            short = re.sub(r"<[^>]*>", "", fn.split("(")[0])
            loc = m.group(2).replace("/home/zhuhe/code/litchi/","")
            loc = loc + (":" + m.group(3) if m.group(3) else "")
            cur.append(f"{short} [{loc.split('/')[-1]}]")
print(f"== {p.stem}")
for ph, c in out.items():
    if ph in ("pre","done"): continue
    print(f"  phase {ph}: {sum(c.values())} positional reads")
    for chain, n in c.most_common():
        print(f"    x{n:3d}  {chain}")
