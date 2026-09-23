#!/usr/bin/env python3
"""Change 0754 worst-case witnesses for namespace resolution: a DOCX main part
with about 30,000 in-scope declarations (120 nested elements declaring 255
prefixes each, inside the text scanner's 128-deep limit), followed by 20,000
elements whose names resolve against that list. Variants: every name uses one
distinct undeclared prefix (each lookup misses: the linear search's worst
case), and every name uses the outermost declared prefix (base searches the
whole list for each; a cached position answers the rest). The `-tab`
variants use a local name the DOCX text scan always resolves."""
import sys, zipfile
template, out = sys.argv[1], sys.argv[2]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def document(names):
    parts = [f'<w:document xmlns:w="{W}"><w:body><w:p><w:r>']
    for level in range(120):
        declarations = "".join(f' xmlns:p{level}x{k}="urn:{level}:{k}"' for k in range(255))
        parts.append(f"<w:x{declarations}>")
    parts.append("<w:t>text</w:t>")
    parts.extend(names)
    parts.append("</w:x>" * 120)
    parts.append("</w:r></w:p></w:body></w:document>")
    return "".join(parts).encode()
variants = {
    "dos-miss": [f"<q{i}:e/>" for i in range(20000)],
    "dos-outer": ["<p0x0:e/>" for _ in range(20000)],
    "dos-default": ["<e/>" for _ in range(20000)],
    # Names the text scan must resolve (a special character's local name).
    "dos-miss-tab": [f"<q{i}:tab/>" for i in range(20000)],
    "dos-outer-tab": ["<p0x0:tab/>" for _ in range(20000)],
}
z = zipfile.ZipFile(template)
for name, names in variants.items():
    xml = document(names)
    with zipfile.ZipFile(f"{out}/{name}.docx", "w", zipfile.ZIP_DEFLATED) as target:
        for info in z.infolist():
            data = xml if info.filename == "word/document.xml" else z.read(info.filename)
            target.writestr(info.filename, data)
    print(name, len(xml), "bytes")
