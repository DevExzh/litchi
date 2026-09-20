#!/usr/bin/env python3
"""Reproduce 0649's same-length MCE-marker counterfactual; not a valid edit oracle."""
import hashlib
import json
import zipfile
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SOURCE = ROOT / 'test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx'
OLD = b'http://schemas.openxmlformats.org/markup-compatibility/2006'
NEW = b'http://schemas.openxmlformats.org/markup-kompatibility/2006'
assert len(OLD) == len(NEW)
output = P / 'marker-control.pptx'
changes = []
with zipfile.ZipFile(SOURCE) as source, zipfile.ZipFile(output, 'w') as dest:
    for info in source.infolist():
        before = source.read(info)
        after = before.replace(OLD, NEW)
        dest.writestr(info, after)
        changes.append(dict(member=info.filename, bytes=len(before), replacements=before.count(OLD),
                            before_sha256=hashlib.sha256(before).hexdigest(),
                            after_sha256=hashlib.sha256(after).hexdigest()))
manifest = dict(source=str(SOURCE.relative_to(ROOT)),
                source_sha256=hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
                control=output.name, control_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),
                note='Changes the namespace URI to bypass MCE detection. This is a mechanism control, not semantic equivalence; ZIP compression is regenerated.',
                members=changes)
(P / 'control-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
print(output)
