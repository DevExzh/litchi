"""Recompute report JSON and Markdown into owned scratch and compare exact bytes."""
import json
import shutil
import subprocess
from run import HERE, REPO, SCRATCH, sha


def main():
    directory = SCRATCH / 'report-replay'
    directory.mkdir()
    records = []
    jobs = [
        ('baseline-native-analysis', 'analyze_native.py', []),
        ('candidate-native-analysis', 'analyze_native.py', ['--lanes', 'after-r1', 'after-r2']),
        ('candidate-comparison', 'compare_candidate.py', []),
        ('profile-comparison', 'compare_profile_lanes.py', []),
        ('publication-allocation-comparison', 'analyze_allocations.py', []),
    ]
    for name, script, extra in jobs:
        command = ['python3', '-B', str(HERE / script), *extra, ('--output' if script == 'analyze_allocations.py' else '--json'), str(directory / (name + '.json')), '--markdown', str(directory / (name + '.md'))]
        result = subprocess.run(command, cwd=REPO, capture_output=True, text=True)
        assert result.returncode == 0, (name, result.stderr)
        hashes = {}
        for suffix in ['.json', '.md']:
            filename = name + suffix
            assert (directory / filename).read_bytes() == (HERE / filename).read_bytes(), filename
            hashes[filename] = sha(directory / filename)
        records.append(dict(command=command, exit_code=result.returncode, helper_sha256=sha(HERE/script), output_sha256=hashes))
        print(name, 'exact replay passed', flush=True)
    shutil.rmtree(directory)
    (HERE / 'report-replay.json').write_text(json.dumps(dict(reports=records, temporary_replay_removed=not directory.exists()), indent=2)+'\n')

if __name__ == '__main__':
    main()
