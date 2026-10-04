"""Offline replay and exact retained-benchmark identity check for change 0904."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile

P = Path(__file__).resolve().parent

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def verify():
    prior = json.loads((P / 'prior-identities.json').read_text())
    manifest = json.loads((P / 'artifacts/manifest.json').read_text())
    assert {case['id'] for case in manifest['cases']} == set(prior)
    rows = []
    for case in manifest['cases']:
        old = prior[case['id']]
        report = Path(old['report']['path'])
        if not report.is_absolute():
            report = P.parents[3] / report
        assert report.stat().st_size == old['report']['bytes']
        assert digest(report) == old['report']['sha256']
        raw = json.loads(report.read_text())['results'][0]
        assert raw['corpus'] == old['corpus']
        assert raw['output_sha256'] == old['output_sha256']
        assert case['source']['sha256'] == raw['corpus']['archive_sha256']
        assert case['output']['sha256'] == raw['output_sha256']
        with zipfile.ZipFile(P/'artifacts'/case['source']['path']) as source, zipfile.ZipFile(P/'artifacts'/case['output']['path']) as output:
            assert set(source.namelist()) == set(output.namelist())
            changed = sorted(name for name in source.namelist() if source.read(name) != output.read(name))
        assert changed == sorted(part['member'] for part in case['parts'])
        rows.append({'case':case['id'], 'prior_archive_identity':True, 'selected_changed_members':changed})
    with tempfile.TemporaryDirectory(prefix='litchi-0904-replay-') as directory:
        replay = Path(directory)/'analysis.json'
        command = [sys.executable, '-B', str(P/'census.py'), str(P/'artifacts'), '--output', str(replay)]
        subprocess.run(command, check=True, capture_output=True)
        assert replay.read_bytes() == (P/'analysis.json').read_bytes()
        before = digest(replay)
        refused = subprocess.run(command, capture_output=True)
        assert refused.returncode != 0 and digest(replay) == before
    return {'status':'pass', 'cases':rows, 'canonical_replay_byte_identical':True, 'census_existing_output_refused_without_mutation':True, 'temporary_replay_removed':True, 'analysis_sha256':digest(P/'analysis.json')}

if __name__ == '__main__':
    print(json.dumps(verify(), indent=2, sort_keys=True))
