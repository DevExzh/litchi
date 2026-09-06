#!/usr/bin/env python3
"""Finish the interrupted supplementary profiles without rerecording completed CPU data."""
import datetime
import hashlib
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    build = json.loads((ROOT / 'build.json').read_text())
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert digest(Path(build['capture_binary'])) == build['binary_sha256']
    assert digest(Path(__file__)) == build['additional_artifact_sha256']['resume-profiles.py']
    assert digest(ROOT / 'verify-report.py') == build['verifier_sha256']
    assert digest(ROOT / 'protocol.json') == build['protocol_sha256']
    plan = protocol['supplementary_profiles']
    target = ROOT / 'profiles'
    assert target.is_dir() and not (ROOT / 'profile-index.json').exists()
    receipts = []
    for index, case in enumerate(plan['cases']):
        name = case['corpus'] + '-' + case['provider']
        data, report, log = (target / (name + suffix) for suffix in ['.data', '.json', '.log'])
        receipt_path = target / (name + '-receipt.json')
        assert not receipt_path.exists()
        argv = ['taskset', '-c', str(protocol['cpu']), 'perf', 'record', '-e', plan['event'],
                '-F', str(plan['frequency_hz']), '--call-graph', plan['call_graph'], '-o', str(data), '--',
                build['capture_binary'], case['command'], '--corpus', case['corpus'], '--provider', case['provider'],
                '--samples', str(plan['samples']), '--warmup', str(plan['warmups']),
                '--source-revision', build['revision'], '--output', str(report)]
        recovered = index == 0
        receipt = {'name': name, 'argv': argv, 'revision': build['revision'], 'binary_sha256': build['binary_sha256'],
                   'scope': plan['scope'], 'recovered_completed_recording': recovered,
                   'recording_started_utc': None if recovered else now(), 'postprocessing_started_utc': now()}
        if recovered:
            assert all(path.is_file() for path in [data, report, log])
            assert 'perf record: Captured and wrote' in log.read_text()
            receipt['recovery_source'] = 'checks/release-profiles.json'
            receipt['argv_provenance'] = 'Reconstructed from the frozen profiles.py and protocol; original wrapper stopped after successful perf record at the sample-count validator.'
            result_code = 0
        else:
            assert not any(path.exists() for path in [data, report, log])
            print('PROFILE ' + name, flush=True)
            with log.open('xb') as output:
                result_code = subprocess.run(argv, cwd=REPO, stdout=output, stderr=subprocess.STDOUT).returncode
        receipt.update(exit_code=result_code, status='recorded' if result_code == 0 else 'failed')
        receipt_path.write_text(json.dumps(receipt, indent=2) + '\n')
        assert result_code == 0
        subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(report)], check=True)
        row = json.loads(report.read_text())
        assert row['binary_sha256'] == build['binary_sha256'] and row['source_revision'] == build['revision']
        assert row['corpus'] == case['corpus'] and row['provider'] == case['provider']
        assert row['samples'] == plan['samples'] and row['warmup'] == plan['warmups']
        for suffix, command in [('script.log', ['perf', 'script', '-i', str(data)]),
                                ('report.log', ['perf', 'report', '--stdio', '--no-children', '-i', str(data)])]:
            with (target / (name + '-' + suffix)).open('xb') as output:
                subprocess.run(command, cwd=REPO, stdout=output, stderr=subprocess.STDOUT, check=True)
        receipt['artifacts'] = {str(path.relative_to(ROOT)): {'sha256': digest(path), 'bytes': path.stat().st_size}
                                for path in [data, report, log, target / (name + '-script.log'), target / (name + '-report.log')]}
        receipt.update(status='pass', finished_utc=now())
        receipt_path.write_text(json.dumps(receipt, indent=2) + '\n')
        receipts.append(str(receipt_path.relative_to(ROOT)))
    (ROOT / 'profile-index.json').write_text(json.dumps(receipts, indent=2) + '\n')


if __name__ == '__main__':
    main()
