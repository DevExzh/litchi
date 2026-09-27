"""Current-source descriptive OPC publication captures; no historical ratio."""
import json
import pathlib
import subprocess
import time

from build import P, ROOT, guard, sha


if __name__ == '__main__':
    guard()
    build = json.loads((P / 'build.json').read_text())
    binary = pathlib.Path(build['binary'])
    assert sha(binary) == build['binary_sha256']
    folder = P / 'captures'
    folder.mkdir()
    rows = []
    expected = None
    for repeat in range(3):
        report = folder / f'{repeat:02d}.json'
        command = ['taskset', '-c', '12', str(binary), '--warmup', '2', '--samples', '9',
                   '--case', 'opc_mutated_save,opc_noop_save', '--json', str(report)]
        row = {'repeat': repeat, 'command': command, 'started': time.time(),
               'monotonic_start': time.monotonic_ns(), 'build_sha256': sha(P / 'build.json')}
        with report.with_suffix('.stdout').open('w') as out, report.with_suffix('.stderr').open('w') as err:
            process = subprocess.run(command, cwd=ROOT, stdout=out, stderr=err)
        row |= {'exit': process.returncode, 'ended': time.time(), 'monotonic_end': time.monotonic_ns()}
        row['files'] = {str(f.relative_to(P)): sha(f) for f in
                        (report, report.with_suffix('.stdout'), report.with_suffix('.stderr')) if f.is_file()}
        rows.append(row)
        (folder / 'manifest.json').write_text(json.dumps(rows, indent=2) + '\n')
        assert process.returncode == 0, row
        data = json.loads(report.read_text())
        assert data['binary_identity']['binary_sha256'] == build['binary_sha256']
        assert data['environment']['cpu_affinity'] == '12'
        results = data['results']
        assert len(results) == 16
        projection = {}
        for result in results:
            key = result['case'] + '/' + result['corpus']['name']
            assert key not in projection
            assert len(result['elapsed_ns']['samples']) == 9
            assert sorted(result['elapsed_ns']['sample_order']) == list(range(9))
            assert all(type(n) is int and n > 0 for n in result['elapsed_ns']['samples'])
            projection[key] = {'corpus': result['corpus'], 'sink': result['sink'],
                               'output_sha256': result.get('output_sha256')}
        if expected is None:
            expected = projection
            (P / 'descriptive-projection.json').write_text(json.dumps(expected, indent=2) + '\n')
        assert projection == expected
        guard()
        assert sha(binary) == build['binary_sha256']
        print(f'PASS descriptive process {repeat + 1}/3, 16 cases', flush=True)
