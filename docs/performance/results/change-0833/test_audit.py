"""Mutation checks against retained cold proof, output and sample evidence."""
import copy
import json
from pathlib import Path
import sys
import tempfile
import audit
import driver as d

P=d.P

def run():
    tests=[]
    original=d.read(P/'qualification-03.json')
    cases=[
      ('cold-status',lambda r:r['filesystem_evidence'][0]['samples'][1]['cold_verified'].__setitem__('status','ineligible_source_resident')),
      ('cold-resident',lambda r:r['filesystem_evidence'][0]['samples'][1]['cold_verified'].__setitem__('resident_bytes',4096)),
      ('cold-read-bytes',lambda r:r['filesystem_evidence'][0]['samples'][1]['process_metrics'].__setitem__('read_bytes',0)),
      ('cold-post-tool',lambda r:r['filesystem_evidence'][0]['samples'][1]['cold_verified']['fincore_post'].__setitem__('fincore_sha256','0'*64)),
      ('archive-identity',lambda r:r['filesystem_evidence'][0]['corpus'].__setitem__('archive_sha256','0'*64)),
      ('output-identity',lambda r:r['filesystem_evidence'][0]['samples'][1].__setitem__('output_sha256',audit.OUTPUT)),
      ('output-padding',lambda r:r['filesystem_evidence'][0]['samples'][1].__setitem__('output_bytes',16783632)),
      ('read-count',lambda r:r['filesystem_evidence'][0]['samples'][0].__setitem__('logical_read_calls',0)),
      ('sample-clock',lambda r:r['results'][0]['elapsed_ns'].__setitem__('samples',[1])),
      ('binary-identity',lambda r:r['binary_identity'].__setitem__('binary_sha256','0'*64)),
      ('missing-state',lambda r:r['filesystem_evidence'][0]['samples'].pop()),
      ('configuration',lambda r:r['configuration'].__setitem__('cases',[d.CASES[0]])),
    ]
    with tempfile.TemporaryDirectory(prefix='litchi-0833-audit-') as tmp:
        path=Path(tmp)/'report.json'
        path.write_text(json.dumps(original))
        assert audit.report(path,d.CASES[3],['warm','cold-verified'])==2
        for name,mutate in cases:
            value=copy.deepcopy(original);mutate(value);path.write_text(json.dumps(value))
            try: audit.report(path,d.CASES[3],['warm','cold-verified'])
            except (AssertionError,KeyError,StopIteration): tests.append(name)
            else: raise AssertionError(f'mutation accepted: {name}')
    return dict(status='pass',valid_control=1,rejected_mutations=tests,
                audit_sha256=d.sha(P/'audit.py'),script_sha256=d.sha(P/'test_audit.py'))

if __name__=='__main__':
    v=run()
    if sys.argv[1:]==['--write']:d.write(P/'audit-tests.json',v)
    else:
        assert not sys.argv[1:];assert d.read(P/'audit-tests.json')==v
    print('Audit mutation checks PASS: valid control plus 12 rejected mutations')
