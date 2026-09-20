#!/usr/bin/env python3
"""Require successful final-source verification before freezing capture inputs."""
import hashlib
import json
from pathlib import Path
P=Path(__file__).resolve().parent
def read(n):return json.loads((P/n).read_text())
def sha(n):return hashlib.sha256((P/n).read_bytes()).hexdigest()
def main():
    assert not (P/'capture-freeze.json').exists()
    quality=read('quality.json');evidence=read('evidence/results.json')
    assert len(quality)==6 and len(evidence)==6
    assert all(r['exit_code']==0 for r in quality+evidence)
    assert read('source-final.json')==read('source.json')
    for row in quality:
        assert row['source_manifest_sha256']==sha('source.json')
        assert row['log_sha256']==sha(row['log'])
    for row in evidence:
        assert row['source_manifest_sha256']==sha('source-final.json')
        assert row['log_sha256']==sha('evidence/'+row['name']+'.log')
    names=['plan.json','capture.py','custody.py','build.py','builds.json','source.json','source-baseline.json','source-preparation.json','source.patch','constraints.json','helper-freeze.json','protocol-review.json','quality.json','support-files.json','evidence/results.json','freeze-capture.py']
    (P/'capture-freeze.json').write_text(json.dumps({n:sha(n) for n in names},indent=2)+'\n')
    print('Capture inputs frozen after six quality checks and six repository gates.')
if __name__=='__main__':main()
