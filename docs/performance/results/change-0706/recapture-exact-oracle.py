#!/usr/bin/env python3
"""One-time, recorded correction from normalized to exact public Debug outcomes."""
import json
from pathlib import Path
import shutil
import subprocess
from build import ROOT, census, sha

P = Path(__file__).resolve().parent
archive = P/'normalized-oracle-preflight'
assert not archive.exists()
archive.mkdir()
probe = P/'oracle-probe/src/main.rs'
shutil.copy2(probe, archive/'main.rs')
for pattern in ['oracle-build-baseline.*','oracle-build-candidate.*',
                'oracle-baseline-differential*','oracle-candidate-differential*','oracle-baseline-A1*']:
    for path in P.glob(pattern):
        shutil.move(path, archive/path.name)
source = probe.read_text()
source = source.replace('normalize_debug(&format!("{result:?}"))', 'format!("{result:?}")')
start = source.index('fn normalize_debug(')
end = source.index('fn limits_for(', start)
source = source[:start]+source[end:]
source = source.replace('complete normalized Debug result', 'complete exact Debug result')
source = source.replace('every normalized public result', 'every exact public result')
probe.write_text(source)
production = ROOT/'crates/xml-minifier/src/audit.rs'
tests = ROOT/'crates/xml-minifier/tests/attribute_cardinality.rs'
candidate_bytes, test_bytes = production.read_bytes(), tests.read_bytes()
before = census()
revision = json.loads((P/'plan.json').read_text())['revision']
baseline_bytes = subprocess.check_output(['git','show',revision+':crates/xml-minifier/src/audit.rs'],cwd=ROOT)
def run(action, role):
    subprocess.run(['python3','-B',str(P/'oracle.py'),action,role],cwd=ROOT,check=True)
try:
    production.write_bytes(baseline_bytes)
    tests.unlink()
    assert census() == json.loads((P/'source-baseline.json').read_text())
    run('build','baseline')
    run('differential','baseline')
    run('A1','baseline')
finally:
    production.write_bytes(candidate_bytes)
    tests.write_bytes(test_bytes)
assert census() == before
run('build','candidate')
run('differential','candidate')
left = json.loads((P/'oracle-baseline-differential.json').read_text())
right = json.loads((P/'oracle-candidate-differential.json').read_text())
assert left == right
assert all(result['debug'] != 'PANIC' for outcome in left['outcomes'] for result in outcome['results'])
(P/'oracle-exact-correction.json').write_text(json.dumps(dict(
    reason='Independent review found whitespace normalization could hide changes to error details',
    correction='Retain exact Debug strings; recapture both builds and baseline microguard before candidate timings',
    archived='normalized-oracle-preflight', source_restored_exactly=True,
    source_manifest_sha256=sha(P/'source-candidate.json'),
    baseline_source_verified_exactly=True, exact_full_json_equal=True,
    panics=0, casecount=left['casecount'], callcount=left['callcount'],
    script_sha256=sha(Path(__file__))),indent=2)+'\n')
print('Exact oracle parity passed',left['casecount'],left['callcount'],flush=True)
