#!/usr/bin/env python3
"""Bind both source modes to the same current executables, sources and protocol."""
import hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
receipt=json.loads((ROOT/'checks/after-build.json').read_text());assert receipt['status']=='pass' and receipt['source_unchanged']
def sha(name):return hashlib.sha256((ROOT/name).read_bytes()).hexdigest()
for role in ('observed','plain'):
    directory=ROOT/role;directory.mkdir()
    row={'change':445,'role':role,'revision':receipt['revision'],'source_manifest':receipt['source_after'],'binaries':json.loads((ROOT/'after/binary-copies.json').read_text()),'protocol_sha256':sha('protocol.json'),'capture_driver_sha256':sha('capture.py'),'profile_driver_sha256':sha('profile.py'),'build_receipt':'checks/after-build.json'}
    with (directory/'build.json').open('x') as stream:stream.write(json.dumps(row,indent=2)+'\n')
print('Bound observed and plain to one build')
