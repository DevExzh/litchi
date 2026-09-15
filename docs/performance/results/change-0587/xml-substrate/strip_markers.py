"""Make a control variant of an xlsx with the real-producer worksheet markers
removed (x14ac:dyDescent attributes, xmlns:mc, mc:Ignorable, xmlns:x14ac).
Everything else is byte-identical. Survey-only scratch tool."""
import re
import sys
import zipfile

src, dst = sys.argv[1], sys.argv[2]
pat = re.compile(rb' (?:x14ac:dyDescent|xmlns:mc|mc:Ignorable|xmlns:x14ac)="[^"]*"')
stats = {}
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
            new, n = pat.subn(b"", data)
            stats[item.filename] = (len(data), len(new), n)
            data = new
        zout.writestr(item, data)
for name, (before, after, n) in stats.items():
    print(f"{name}: {before} -> {after} bytes, {n} attributes removed")
