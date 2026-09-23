#!/usr/bin/env python3
"""Change 0754 differential inputs: every DOCX/DOTX/DOCM and PPTX fixture under
test-data, the three semantic corpora, and deterministic mutations of each DOCX
main part (byte-level damage and namespace/structure insertions), written as
DOCX files and as raw main-part XML for the snapshot scanner."""
import os, random, sys, zipfile

repo, corpus, out = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(f"{out}/docx", exist_ok=True)
os.makedirs(f"{out}/xml", exist_ok=True)
rng = random.Random(0x0754)

fixtures, pptx = [], []
for root, _dirs, files in os.walk(os.path.join(repo, "test-data")):
    for name in files:
        low = name.lower()
        path = os.path.join(root, name)
        if low.endswith((".docx", ".dotx", ".docm", ".dotm")):
            fixtures.append(path)
        elif low.endswith((".pptx", ".potx", ".pptm")):
            pptx.append(path)
fixtures.sort(); pptx.sort()
fixtures += [os.path.join(corpus, f"semantic-{s}.docx") for s in ("tiny", "medium", "large")]

W = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
INSERTIONS = [
    b'<w:p xmlns:w="urn:foreign"><w:r><w:t>shadowed</w:t></w:r></w:p>',
    b'<w:p><w:r xmlns:w=""><w:t>undeclared</w:t></w:r></w:p>',
    b'<w:p><w:r><x:y/><w:t>undeclared prefix</w:t></w:r></w:p>',
    b'<w:p><w:r><w:t>a]]>b</w:t></w:r></w:p>',
    b'<w:p><w:r><w:t xml:space="sometimes">odd space</w:t></w:r></w:p>',
    b'<w:p xmlns:q="' + W + b'"><q:r><q:t>other prefix</q:t></q:r></w:p>',
    b'<w:p><w:r><w:t>one</w:t><w:tab/><w:br/><w:t>two</w:t></w:r></w:p>',
    b'<w:p xmlns=""><w:r><w:t>default unset</w:t></w:r></w:p>',
    b'<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    b'<w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>',
    b'<w:sectPr/>',
    b'<w:body/>',
    b'<?pi x?>',
    b'<!--c--><w:p/>',
    b'<w:p><w:r><w:t>&bogus;</w:t></w:r></w:p>',
    b'<w:p xmlns:xml="urn:bad"/>',
]
BYTES = b'<>/="\':x \x00&;!?-[]w'

def mutate(xml):
    kind = rng.randrange(7)
    data = bytearray(xml)
    if kind == 0 and data:
        data[rng.randrange(len(data))] = BYTES[rng.randrange(len(BYTES))]
    elif kind == 1 and data:
        at = rng.randrange(len(data)); del data[at:at + 1 + rng.randrange(24)]
    elif kind == 2 and data:
        data = data[:rng.randrange(len(data))]
    else:
        body = data.find(b"<w:body>")
        insertion = INSERTIONS[rng.randrange(len(INSERTIONS))]
        if body >= 0 and kind != 6:
            at = body + len(b"<w:body>")
            # sometimes insert later in the body, between two paragraphs
            later = data.find(b"</w:p>", at + rng.randrange(max(1, len(data) - at)))
            if later >= 0 and rng.randrange(2):
                at = later + len(b"</w:p>")
            data[at:at] = insertion
        elif data:
            at = rng.randrange(len(data)); data[at:at] = insertion
    return bytes(data)

docx_list, xml_list = [], []
for index, path in enumerate(fixtures):
    try:
        z = zipfile.ZipFile(path)
        main = z.read("word/document.xml")
    except Exception:
        docx_list.append(path)
        continue
    docx_list.append(path)
    base_xml = f"{out}/xml/f{index:03d}.xml"
    open(base_xml, "wb").write(main); xml_list.append(base_xml)
    count = 3 if len(main) > 400_000 else 10
    for m in range(count):
        mutated = mutate(main)
        if rng.randrange(3) == 0:
            mutated = mutate(mutated)
        name = f"f{index:03d}-m{m:02d}"
        xml_path = f"{out}/xml/{name}.xml"
        open(xml_path, "wb").write(mutated); xml_list.append(xml_path)
        docx_path = f"{out}/docx/{name}.docx"
        with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as target:
            for info in z.infolist():
                data = mutated if info.filename == "word/document.xml" else z.read(info.filename)
                target.writestr(info.filename, data)
        docx_list.append(docx_path)

open(f"{out}/docx-list.txt", "w").write("\n".join(docx_list) + "\n")
open(f"{out}/xml-list.txt", "w").write("\n".join(xml_list) + "\n")
open(f"{out}/pptx-list.txt", "w").write("\n".join(pptx) + "\n")
print(len(fixtures), "docx fixtures;", len(docx_list), "docx inputs;", len(xml_list), "xml inputs;", len(pptx), "pptx inputs")
