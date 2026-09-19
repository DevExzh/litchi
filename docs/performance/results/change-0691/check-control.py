#!/usr/bin/env python3
"""Verify the marker counterfactual and disclose ZIP metadata differences."""
import hashlib
import json
import zipfile
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
manifest = json.loads((P/'control-manifest.json').read_text())
fields = ['date_time','compress_type','create_system','external_attr','internal_attr',
          'extra','comment','create_version','extract_version','flag_bits']
changes = {}
old = b'http://schemas.openxmlformats.org/markup-compatibility/2006'
new = b'http://schemas.openxmlformats.org/markup-kompatibility/2006'
with zipfile.ZipFile(ROOT/manifest['source']) as source, zipfile.ZipFile(P/manifest['control']) as control:
    assert source.namelist()==control.namelist()
    for row in manifest['members']:
        before,after = source.read(row['member']),control.read(row['member'])
        assert len(before)==len(after)==row['bytes']
        assert before.replace(old,new)==after
        assert before.count(old)==row['replacements']
        assert hashlib.sha256(before).hexdigest()==row['before_sha256']
        assert hashlib.sha256(after).hexdigest()==row['after_sha256']
    for field in fields:
        names = [i.filename for i in source.infolist()
                 if getattr(i,field)!=getattr(control.getinfo(i.filename),field)]
        if names:
            changes[field]=names
    archive_comment_equal = source.comment==control.comment
(P/'control-metadata-check.json').write_text(json.dumps(dict(fields=fields,changed=changes,
    archive_comment_equal=archive_comment_equal),indent=2)+'\n')
print('PASS marker/member invariants; metadata changes:', {k:len(v) for k,v in changes.items()})
