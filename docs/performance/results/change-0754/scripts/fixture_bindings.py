#!/usr/bin/env python3
"""Change 0754 review: the peak number of namespace bindings in scope (the
two reserved ones included, as the tracker counts them) in each DOCX
fixture's main part, and in the main part of every other OOXML fixture the
DOCX text and layout scans do not read, for context."""
import os, sys, zipfile, xml.parsers.expat
repo = sys.argv[1]
rows = []
for root, _dirs, files in os.walk(os.path.join(repo, "test-data")):
    for name in files:
        if not name.lower().endswith((".docx", ".docm", ".dotx", ".dotm")):
            continue
        path = os.path.join(root, name)
        try:
            data = zipfile.ZipFile(path).read("word/document.xml")
        except Exception:
            continue
        stack, peak, current = [], 2, 2
        def start(tag, attrs):
            global current, peak
            count = sum(1 for key in attrs if key == "xmlns" or key.startswith("xmlns:"))
            stack.append(count)
            current += count
            peak = max(peak, current)
        def end(tag):
            global current
            current -= stack.pop()
        parser = xml.parsers.expat.ParserCreate(namespace_separator=None)
        parser.ordered_attributes = False
        parser.StartElementHandler = start
        parser.EndElementHandler = end
        try:
            parser.Parse(data, True)
        except Exception as error:
            rows.append((os.path.relpath(path, repo), None, str(error)))
            continue
        rows.append((os.path.relpath(path, repo), peak, ""))
rows.sort(key=lambda row: (-(row[1] or 0), row[0]))
for path, peak, error in rows:
    print(f"{peak if peak is not None else '-':>4}  {path}  {error}")
print("fixtures:", len(rows), "max peak:", max(r[1] or 0 for r in rows), "over 64:", sum(1 for r in rows if (r[1] or 0) > 64), "over 32:", sum(1 for r in rows if (r[1] or 0) > 32))
