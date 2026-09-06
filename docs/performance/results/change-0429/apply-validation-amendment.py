#!/usr/bin/env python3
"""Bind the explicit profile-validator correction without changing captured data."""
import copy
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def record(path):
    return {'sha256': digest(path), 'bytes': path.stat().st_size}


def main():
    target = ROOT / 'validation-amendment.json'
    assert not target.exists()
    original_dir = ROOT / 'validation-amendment-original'
    original = json.loads((original_dir / 'build.json').read_text())
    build = json.loads((ROOT / 'build.json').read_text())
    assert build == original
    amended = copy.deepcopy(build)
    amended['verifier_sha256'] = digest(ROOT / 'verify-report.py')
    amended['replay_verifier_sha256'] = digest(ROOT / 'verify.py')
    amended['planned_checks_sha256'] = digest(ROOT / 'planned-checks.json')
    amended['additional_artifact_sha256']['resume-profiles.py'] = digest(ROOT / 'resume-profiles.py')
    paths = sorted((ROOT / 'capture').iterdir())
    assert len(paths) == 128 and all(path.is_file() for path in paths)
    paths += [ROOT / 'profiles' / ('media-rich-bytes' + suffix) for suffix in ['.data', '.json', '.log']]
    amendment = {
        'reason': 'The frozen 32-process baseline validated successfully. The frozen report validator omitted the predeclared 100-sample supplementary profile policy. The correction permits (100,3) only for media-rich bytes/file profiles; baseline (30,3) and controls (1,0) remain unchanged. Existing bytes CPU data is postprocessed without rerecording; file CPU capture continues separately.',
        'original_artifacts': {str(path.relative_to(ROOT)): record(path) for path in sorted(original_dir.iterdir())},
        'unchanged_capture_artifacts': {str(path.relative_to(ROOT)): record(path) for path in paths},
        'changed_build_fields': ['verifier_sha256', 'replay_verifier_sha256', 'planned_checks_sha256', 'additional_artifact_sha256', 'validation_amendment'],
        'source_revision': original['revision'], 'binary_sha256': original['binary_sha256'],
    }
    target.write_text(json.dumps(amendment, indent=2) + '\n')
    amended['validation_amendment'] = {'path': target.name, 'sha256': digest(target)}
    (ROOT / 'build.json').write_text(json.dumps(amended, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'unchanged_capture_artifacts': len(paths), 'baseline_recaptured': False}))


if __name__ == '__main__':
    main()
