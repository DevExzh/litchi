"""Replay frozen-source identity and descriptive analysis after temporary cleanup."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def sha(data):
    return hashlib.sha256(data).hexdigest()

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    assert not args.output.exists()
    source = json.loads((P / 'source.json').read_text())
    archive = subprocess.check_output(['git', 'archive', source['commit'], '--', *source['archive_paths']], cwd=ROOT)
    assert sha(archive) == source['archive_sha256']
    files = {}
    with tarfile.open(fileobj=io.BytesIO(archive)) as tar:
        for entry in tar:
            if entry.isfile():
                data = tar.extractfile(entry).read()
                files[entry.name] = {'path': entry.name, 'bytes': len(data), 'sha256': sha(data)}
            elif not entry.isdir():
                raise AssertionError('unexpected non-file archive entry: ' + entry.name)
    for name, overlay in source['lock_overlays'].items():
        data = (ROOT / overlay['artifact']).read_bytes()
        assert sha(data) == overlay['sha256']
        files[name] = {'path': name, 'bytes': len(data), 'sha256': sha(data)}
    marker = (source['commit'] + '\n').encode()
    files['.owned-0907'] = {'path': '.owned-0907', 'bytes': len(marker), 'sha256': sha(marker)}
    assert files == {row['path']: row for row in source['files']}
    plan = json.loads((P / 'capture-plan.json').read_text())
    previous_finish = None
    for label, _ in plan['legs']:
        receipt = json.loads((P / (label + '-run.json')).read_text())
        assert receipt['exit_code'] == 0 and receipt['binary'] == plan['binary']
        assert receipt['environment'] == plan['environment']
        assert previous_finish is None or previous_finish <= receipt['started_utc']
        previous_finish = receipt['finished_utc']
        for key in ['result', 'stdout', 'stderr']:
            descriptor = receipt[key]
            data = (P / Path(descriptor['path']).name).read_bytes()
            assert len(data) == descriptor['bytes'] and sha(data) == descriptor['sha256']
    with tempfile.TemporaryDirectory(prefix='litchi-0907-replay-') as tmp:
        output = Path(tmp) / 'analysis.json'
        subprocess.run(['python3', '-B', str(P / 'analyze.py'), '--packet', str(P), '--output', str(output)], check=True)
        # Analyzer artifact paths must be stable across the temporary output path.
        assert json.loads(output.read_text()) == json.loads((P / 'analysis.json').read_text())
    result = {'status': 'pass', 'source_files_verified': len(files), 'source_commit': source['commit'], 'raw_leg_receipts_verified': len(plan['legs']), 'analysis_reproduced': True, 'temporary_directory_removed': not Path(tmp).exists(), 'scope': 'Source reconstruction and retained raw evidence replay; no binary rerun or live-root claim.'}
    with args.output.open('x') as out:
        json.dump(result, out, indent=2)
        out.write('\n')
    print('replay pass', flush=True)

if __name__ == '__main__':
    main()
