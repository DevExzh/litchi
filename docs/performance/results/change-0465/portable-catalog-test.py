#!/usr/bin/env python3
"""Keep historical portable evidence independent of a later repository catalog."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def main():
    with tempfile.TemporaryDirectory(prefix='litchi-0465-catalog-replay-') as directory:
        repo = Path(directory) / 'repo'
        copied = repo / 'docs/performance/results/change-0465'
        copied.parent.mkdir(parents=True)
        shutil.copytree(ROOT, copied)
        (repo / 'tools').mkdir()
        (repo / 'tools/perf_compare.py').write_text('# repository detection marker only\n')
        current = repo / 'docs/performance/results/perf-regression-default-manifest-v1.json'
        current.write_text('{"later_repository_catalog": true}\n')
        subprocess.run([sys.executable, '-B', str(copied / 'seal.py')], check=True,
                       capture_output=True, text=True)
        results = {}
        for mode, flags in [('portable', []), ('precleanup', ['--precleanup'])]:
            argv = [sys.executable, '-B', str(copied / 'verify.py'), *flags]
            result = subprocess.run(argv, cwd=repo, capture_output=True, text=True)
            results[mode] = {'argv': argv, 'exit_code': result.returncode,
                             'stdout': result.stdout, 'stderr': result.stderr}
        assert results['portable']['exit_code'] == 0, results['portable']
        assert json.loads(results['portable']['stdout'])['portable'] is True
        assert results['precleanup']['exit_code'] != 0
        assert 'does not equal repository checked artifact' in results['precleanup']['stdout']
    print(json.dumps({'schema': 'litchi-0465-portable-catalog-test-v1', 'status': 'pass',
                      'verifier_sha256': hashlib.sha256((ROOT/'verify.py').read_bytes()).hexdigest(),
                      'later_catalog_ignored_only_in_portable_mode': True,
                      'temporary_directory_absent': not Path(directory).exists(),
                      'results': results}, indent=2))


if __name__ == '__main__':
    main()
