#!/usr/bin/env python3
"""Generate minimal OPC packages that isolate what an open validates.

Every member is STORED, so the physical layout is transparent and the
packages are byte-deterministic. Nothing here is a real Office document;
each one exists to make exactly one open-time decision observable.
"""
import sys, zipfile, os

CT_REL = "application/vnd.openxmlformats-package.relationships+xml"
CT_DOC = ("application/vnd.openxmlformats-officedocument"
          ".wordprocessingml.document.main+xml")
R_OFFICE = ("http://schemas.openxmlformats.org/officeDocument"
            "/2006/relationships/officeDocument")
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

def content_types(overrides):
    o = "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'<Default Extension="rels" ContentType="{CT_REL}"/>'
            f'<Default Extension="xml" ContentType="application/xml"/>'
            f'{o}</Types>')

def rels(entries):
    r = "".join(
        f'<Relationship Id="{i}" Type="{t}" Target="{tg}"/>' for i, t, tg in entries)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{NS_REL}">{r}</Relationships>')

DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>')

ROOT_RELS = rels([("rId1", R_OFFICE, "word/document.xml")])
GOOD_DOC_RELS = rels([])
MALFORMED_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  f'<Relationships xmlns="{NS_REL}"><Relationship Id="rId1" Ty')
DUP_ID_RELS = rels([
    ("rId1", "http://example.invalid/a", "a.xml"),
    ("rId1", "http://example.invalid/b", "b.xml"),
])

def write(path, members):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_STORED) as z:
        for name, data in members:
            zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_STORED
            zi.external_attr = 0o600 << 16
            z.writestr(zi, data)

def base(doc_rels=GOOD_DOC_RELS, extra=(), overrides=()):
    ov = [("/word/document.xml", CT_DOC)] + list(overrides)
    m = [("[Content_Types].xml", content_types(ov)),
         ("_rels/.rels", ROOT_RELS),
         ("word/document.xml", DOC)]
    if doc_rels is not None:
        m.append(("word/_rels/document.xml.rels", doc_rels))
    m.extend(extra)
    return m

def main(outdir):
    os.makedirs(outdir, exist_ok=True)
    cases = {}

    # Control: a valid package that must open.
    cases["c0-valid"] = base()

    # E1: a deep relationship part, never named by the caller, is malformed XML.
    cases["e1-deep-rels-malformed"] = base(doc_rels=MALFORMED_RELS)

    # E2: a deep relationship part carries a duplicate relationship Id.
    cases["e2-deep-rels-duplicate-id"] = base(doc_rels=DUP_ID_RELS)

    # E3: the malformed relationship part belongs to an ORPHAN typed part --
    # one no relationship reaches -- so only the post-walk fallback loop reads it.
    cases["e3-orphan-rels-malformed"] = base(
        overrides=[("/custom/orphan.xml", "application/vnd.example+xml")],
        extra=[("custom/orphan.xml", "<o/>"),
               ("custom/_rels/orphan.xml.rels", MALFORMED_RELS)])

    # E4: an untyped member that NOTHING refers to is archive junk; the open
    # must succeed and report it as a non-part member.
    cases["e4-untyped-unreferenced"] = base(extra=[("junk/thing.bin", b"\x00\x01\x02")])

    # E5: the SAME untyped member, now named by a deep relationship. Admission
    # flips because the relationship closure -- not the member -- decides.
    cases["e5-untyped-referenced"] = base(
        doc_rels=rels([("rId9", "http://example.invalid/x", "/junk/thing.bin")]),
        extra=[("junk/thing.bin", b"\x00\x01\x02")])

    # E6: a second-level relationship part, two hops from the root, is malformed.
    cases["e6-second-hop-rels-malformed"] = base(
        doc_rels=rels([("rId2", "http://example.invalid/hdr", "header1.xml")]),
        overrides=[("/word/header1.xml", "application/vnd.example+xml")],
        extra=[("word/header1.xml", "<h/>"),
               ("word/_rels/header1.xml.rels", MALFORMED_RELS)])

    for name, members in cases.items():
        p = os.path.join(outdir, f"{name}.docx")
        write(p, members)
        print(f"{os.path.getsize(p):7d}  {name}.docx  members={len(members)}")

if __name__ == "__main__":
    main(sys.argv[1])
