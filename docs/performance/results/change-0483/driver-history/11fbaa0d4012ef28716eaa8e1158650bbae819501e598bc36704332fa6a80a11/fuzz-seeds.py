#!/usr/bin/env python3
"""Generate deterministic stored/Deflate DOCX append fuzz seeds."""
import io
import json
import hashlib
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parent
WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
STRICT = 'http://purl.oclc.org/ooxml/wordprocessingml/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
TYPES = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="bin" ContentType="application/octet-stream"/>'
         '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
         '</Types>')


def document(body, namespace=WORD, prefix='w'):
    q = prefix + ':' if prefix else ''
    declaration = 'xmlns:' + prefix if prefix else 'xmlns'
    return f'<{q}document {declaration}="{namespace}"><{q}body>{body}</{q}body></{q}document>'


def package(xml, method, strict=False):
    out = io.BytesIO()
    rel = STRICT_REL if strict else REL
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}" Target="word/document.xml"/>'
            '</Relationships>')
    with zipfile.ZipFile(out, 'w') as archive:
        for name, data in [('[Content_Types].xml', TYPES.encode()),
                           ('_rels/.rels', rels.encode()),
                           ('word/document.xml', xml.encode()),
                           ('word/opaque.bin', bytes(range(256)))]:
            entry = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            entry.compress_type = method
            entry.external_attr = 0o600 << 16
            archive.writestr(entry, data)
    return out.getvalue()


def main():
    seeds = ROOT / 'fuzz' / 'seeds'
    seeds.mkdir(parents=True, exist_ok=False)
    para = '<w:p><w:r><w:t xml:space="preserve"> seed &amp; text </w:t></w:r></w:p>'
    section = '<w:sectPr xmlns:v="urn:vendor" v:flag="keep"><v:opaque a="&quot;"/></w:sectPr>'
    cases = {
        'plain': (document(para), False),
        'empty-body': (document(''), False),
        'empty-paragraph': (document('<w:p/>'), False),
        'section': (document(para + section), False),
        'only-section': (document(section), False),
        'section-nonfinal': (document(section + para), False),
        'section-duplicate': (document(section + section), False),
        'section-nested': (document('<w:p>' + section + '</w:p>'), False),
        'strict': (document(para, STRICT), True),
        'default-namespace': (document('<p><r><t>seed</t></r></p>', prefix=''), False),
        'local-prefix': (f'<a:document xmlns:a="{WORD}"><b:body xmlns:b="{WORD}"><b:p><b:r><b:t>seed</b:t></b:r></b:p></b:body></a:document>', False),
        'long-token': (document('<w:p><w:r><w:t>' + 'x' * 8193 + '</w:t></w:r></w:p>'), False),
        'comment': (document(para + '<!--opaque-->'), False),
        'table': (document('<w:tbl/>'), False),
        'malformed': (document(para)[:-4], False),
    }
    manifest = {}
    for name, (xml, strict) in cases.items():
        for method, suffix in [(zipfile.ZIP_STORED, 'stored'), (zipfile.ZIP_DEFLATED, 'deflate')]:
            data = package(xml, method, strict)
            path = seeds / f'{name}-{suffix}.docx'
            path.write_bytes(data)
            manifest[path.name] = {'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest(),
                                   'main_xml_sha256': hashlib.sha256(xml.encode()).hexdigest()}
    (ROOT / 'fuzz' / 'seed-manifest.json').write_text(json.dumps(manifest, indent=2, sort_keys=True) + '\n')
    print(f'Generated {len(manifest)} deterministic fuzz seeds.')


if __name__ == '__main__':
    main()
