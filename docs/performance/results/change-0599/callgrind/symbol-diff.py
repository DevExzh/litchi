import re, subprocess, sys
def table(path):
    out = subprocess.run(["callgrind_annotate", "--threshold=99.9", "--auto=no", path],
                         capture_output=True, text=True).stdout
    rows = {}
    for line in out.splitlines():
        m = re.match(r"^\s*([\d,]+) \([ \d.]+%\)\s+(.*?)(?: \[.*\])?$", line)
        if m and "PROGRAM TOTALS" not in line:
            rows[m.group(2).strip()] = int(m.group(1).replace(",", ""))
    return rows
a, b = table(sys.argv[1]), table(sys.argv[2])
keys = set(a) | set(b)
delta = sorted(keys, key=lambda k: b.get(k, 0) - a.get(k, 0))
print(f"{'delta Ir (after-before)':>24s}  {'before':>14s} {'after':>14s}  symbol")
for k in delta[:14]:
    print(f"{b.get(k,0)-a.get(k,0):>24,}  {a.get(k,0):>14,} {b.get(k,0):>14,}  {k}")
print("...")
for k in delta[-4:]:
    print(f"{b.get(k,0)-a.get(k,0):>24,}  {a.get(k,0):>14,} {b.get(k,0):>14,}  {k}")
