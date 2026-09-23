#!/usr/bin/env python3
"""Change 0754 review witnesses: the reviewer's namespace worst cases as DOCX
files, plus the raw main parts for the layout scan.

  defaults-distinct   120 nested w:x declaring 255 xmlns="urn:L:K" each, then 20,000 <qN:tab/>
  defaults-redeclare  the same nesting, then 20,000 <w:tab xmlns="u"/>
  emptyprefix-distinct 120 nested w:x declaring 255 xmlns:="urn:L:K" each, then 20,000 <qN:tab/>
  root40-shortlived   40 extra root declarations, then 200,000 <w:tab xmlns:zN="urn:N"/>
  root70-shortlived   70 extra root declarations, then 200,000 <w:tab xmlns:zN="urn:N"/>
  named-distinct      120 nested w:x declaring 255 xmlns:pLxK each, then 20,000 <qN:tab/>
"""
import sys, zipfile
template, out = sys.argv[1], sys.argv[2]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def nested(declaration, names, root_extra=""):
    parts = [f'<w:document xmlns:w="{W}"{root_extra}><w:body><w:p><w:r>']
    for level in range(120):
        parts.append("<w:x" + "".join(declaration(level, k) for k in range(255)) + ">")
    parts.append("<w:t>text</w:t>")
    parts.extend(names)
    parts.append("</w:x>" * 120)
    parts.append("</w:r></w:p></w:body></w:document>")
    return "".join(parts).encode()

def flat(root_declarations, count):
    extra = "".join(f' xmlns:r{k}="urn:root:{k}"' for k in range(root_declarations))
    parts = [f'<w:document xmlns:w="{W}"{extra}><w:body><w:p><w:r><w:t>text</w:t>']
    parts.extend(f'<w:tab xmlns:z{n}="urn:{n}"/>' for n in range(count))
    parts.append("</w:r></w:p></w:body></w:document>")
    return "".join(parts).encode()

distinct = [f"<q{i}:tab/>" for i in range(20000)]
variants = {
    "defaults-distinct": nested(lambda l, k: f' xmlns="urn:{l}:{k}"', distinct),
    "defaults-redeclare": nested(lambda l, k: f' xmlns="urn:{l}:{k}"', ['<w:tab xmlns="u"/>'] * 20000),
    "emptyprefix-distinct": nested(lambda l, k: f' xmlns:="urn:{l}:{k}"', distinct),
    "named-distinct": nested(lambda l, k: f' xmlns:p{l}x{k}="urn:{l}:{k}"', distinct),
    "root40-shortlived": flat(40, 200000),
    "root70-shortlived": flat(70, 200000),
}
z = zipfile.ZipFile(template)
for name, xml in variants.items():
    with zipfile.ZipFile(f"{out}/{name}.docx", "w", zipfile.ZIP_DEFLATED) as target:
        for info in z.infolist():
            data = xml if info.filename == "word/document.xml" else z.read(info.filename)
            target.writestr(info.filename, data)
    open(f"{out}/{name}.xml", "wb").write(xml)
    print(name, len(xml), "bytes")
