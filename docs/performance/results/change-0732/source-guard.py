#!/usr/bin/env python3
"""Offline ordinary-method and exact source/archive custody check."""
import hashlib
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
r = json.loads((P / 'source-review.json').read_text())

def sha(data):
    return hashlib.sha256(data).hexdigest()

for name, row in r['files'].items():
    before = (P / 'source-archive/before' / name).read_bytes()
    after = (P / 'source-archive/after' / name).read_bytes()
    assert sha(before) == row['before_sha256']
    assert sha(after) == row['after_sha256']
    assert after == (ROOT / name).read_bytes()
    assert before == subprocess.check_output(['git', 'show', r['base_head'] + ':' + name], cwd=ROOT)

def ordinary(data):
    start = data.index(b'    pub fn commit(self) -> Result<Commit> {')
    end = data.index(b'\n    }\n', start) + len(b'\n    }\n')
    return data[start:end]

name = 'crates/litchi-ppt/src/slide_order.rs'
a = ordinary((P / 'source-archive/before' / name).read_bytes())
b = ordinary((P / 'source-archive/after' / name).read_bytes())
assert a == b and sha(b) == r['ordinary_commit_sha256']
assert r['ordinary_commit_unchanged'] is True
print('PASS exact base/archive/current custody and byte-identical ordinary commit')
