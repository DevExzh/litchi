"""Join independent replay, exact source custody, chronology, and cleanup."""
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0814'

def read(path):
    return json.loads(path.read_text())

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def artifact(row):
    path = Path(row['path'])
    assert path.is_file() and not path.is_symlink(), path
    assert path.stat().st_size == row['bytes'] and sha(path) == row['sha256'], path
    return path

def interval(rows):
    assert rows
    previous = 0
    for row in rows:
        assert row['exit_code'] == 0
        assert previous <= row['started'] <= row['ended']
        previous = row['ended']
        artifact(row['log'])
    return rows[0]['started'], previous

for script in ('analysis.py', 'root_audit.py', 'offset_audit.py'):
    subprocess.run([sys.executable, '-B', str(P / script), '--check'], cwd=ROOT, check=True)
a, b = read(P / 'analysis.json'), read(P / 'root-audit.json')
assert (a['native']['reports'], a['native']['samples']) == (54, 1620)
assert (a['perf']['summary']['reports'], a['perf']['summary']['samples']) == (2, 200)
assert (b['reports'], b['samples']) == (56, 1820)
assert len(b['paired_p50']) == 6
for row in b['paired_p50']:
    paired = a['native']['summaries'][row['shape']]['paired_ratios'][row['comparison']]['p50']
    assert paired['ratio_median'] == row['ratio']
    assert [paired['bootstrap']['ci_low'], paired['bootstrap']['ci_high']] == row['ci95']
    assert len(paired['by_block']) == 6
    for left, right in zip(paired['by_block'], row['blocks']):
        assert all(left[k] == right[k] for k in ('block', 'before', 'after', 'ratio'))
assert len(a['perf']['reports']) == len(b['frames']) == 2
for report, cross in zip(a['perf']['reports'], b['frames']):
    stack = report['stack']
    assert stack['whole_process_samples'] == cross['whole_process_samples']
    assert stack['owner_qualified_samples'] == cross['owner_samples']
    assert stack['owner_qualified_period'] == cross['owner_period']
    assert stack['qualified_stacks_with_unknown_interior'] == cross['unknown_interior']
    for name, count in stack['self_leaf_samples']:
        assert cross['leaf_counts'][name] == count
    for name, count in stack['inclusive_symbol_samples']:
        assert cross['frame_occurrences'][name] == count

offsets = read(P / 'offset-audit.json')
assert offsets['schema'] == 'litchi.performance.0814.offset-audit.v1'
assert [r['repeat'] for r in offsets['reports']] == [0, 1]
assert [r['repeat'] for r in a['perf']['reports']] == [0, 1]
assert [r['path'] for r in b['frames']] == ['0.frames.gz', '1.frames.gz']
for row, cross in zip(offsets['reports'], b['frames']):
    assert row['owner_samples'] == cross['owner_samples']
    assert sum(v['leaf_samples'] for v in row['symbols'].values()) + row['other_owner_leaf_samples'] == row['owner_samples']
    for values in row['symbols'].values():
        assert values['leaf_samples'] == cross['leaf_counts'].get(values['symbol'], 0)
        assert sum(item['samples'] for item in values['offsets']) == values['leaf_samples']

build = read(P / 'build/build.json')
frozen = read(artifact(build['frozen_inputs']))
for name, digest in frozen['drivers'].items():
    assert sha(P / name) == digest, name
source = read(artifact(build['source']))
assert len(source['files']) == 9196
assert source['files'] == read(P.parent / 'change-0813/build-after/source.json')['files']
names = subprocess.check_output(['git', 'ls-files', '-z', '--', 'crates', 'Cargo.toml',
                                'clippy.toml', '.cargo/config.toml', 'rust-toolchain.toml'], cwd=ROOT).decode().split('\0')
assert set(filter(None, names)) == set(source['files'])
for name, digest in source['files'].items():
    assert sha(ROOT / name) == digest, name
for group in ('architecture', 'unrelated'):
    for name, digest in frozen[group].items():
        assert sha(ROOT / name) == digest, name
for name, digest in frozen['root_inputs'].items():
    assert sha(ROOT / name) == digest, name
assert set(build['binaries']) == {'control', 'profile', 'fp'}
assert [row['variant'] for row in build['rows']] == ['control', 'profile', 'fp']
for row in build['rows']:
    flags = '-C force-frame-pointers=yes' if row['variant'] == 'fp' else None
    features = [] if row['variant'] == 'control' else ['--features', 'capture-profile']
    command = ['cargo', 'build', '--offline', '--locked', '--release', '--manifest-path',
               str(P / 'probe-src/Cargo.toml'), *features]
    assert row['command'] == command and row['features'] == features and row['rustflags'] == flags
    assert row['environment'] == {'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2',
                                  'CARGO_INCREMENTAL': '0', 'RUSTFLAGS': flags}

reuse = read(P / 'quality-reuse.json')
assert reuse['schema'] == 'litchi.performance.0814.quality-reuse.v1'
assert reuse['mode'] == 'reuse-sealed-0813-after-production-gates' and reuse['cargo_executed'] is False
assert reuse['gate_count'] == 6 and reuse['source_files_equal'] is True
old = P.parent / 'change-0813'
old_seal = read(old / 'seal.json')['files']
for descriptor in (reuse['reference'], reuse['source']):
    path = artifact(descriptor)
    assert old_seal[str(path.relative_to(ROOT))] == descriptor['sha256']
assert Path(reuse['reference']['path']) == old / 'quality-summary.json'
assert Path(reuse['source']['path']) == old / 'build-after/source.json'
old_quality = read(old / 'quality-summary.json')['production_after']
assert reuse['gates'] == old_quality['gates'] and reuse['tests'] == old_quality['tests']
for gate in reuse['gates']:
    path = artifact(gate['log'])
    assert old_seal[str(path.relative_to(ROOT))] == gate['log']['sha256']
    assert gate['exit_code'] == 0 and gate['status'] == 'pass'

probe = read(P / 'probe-quality.json')
assert probe['gate_count'] == 3 and probe['tests_passed'] == 36
probe_rows = read(artifact(probe['receipts']))
assert len(probe_rows) == 3
test_log = artifact(probe_rows[1]['log']).read_text()
counts = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;', test_log)
assert sum(int(row[0]) for row in counts) == 36
assert all(int(row[1]) == 0 for row in counts)
lanes = [read(P / 'build/commands.json'), probe_rows, read(P / 'native/receipts.json'),
         read(P / 'perf/receipts.json'), read(P / 'perf/decode-receipts.json')]
assert [len(rows) for rows in lanes] == [3, 3, 54, 2, 2]
spans = [interval(rows) for rows in lanes]
assert all(left[1] <= right[0] for left, right in zip(spans, spans[1:]))

recovery = read(P / 'assembly-recovery.json')
assert recovery['schema'] == 'litchi.performance.0814.assembly-recovery.v1'
assert recovery['observed_exit_code'] == 1 and recovery['workload_rerun'] is False
original_driver = artifact(recovery['failed_driver'])
recovery_driver = artifact(recovery['recovery_driver'])
assert original_driver == P / 'assembly.py' and recovery_driver == P / 'finish_assembly.py'
expected_recovery = original_driver.read_text().replace(
    "'18scan_processed_xml17h'", "'_ZN11litchi_pptx5notes5codec18scan_processed_xml17h'").replace(
    "'15inspect_element17h'", "'_ZN11litchi_pptx5notes5codec15inspect_element17h'").replace(
    'Root-only exact scanner/inspector disassembly before cleanup or workload comparison.',
    'Root-only assembly recovery with fully qualified PPTX symbol selectors.')
assert recovery_driver.read_text() == expected_recovery
assert recovery['binary_witnesses'] == build['binaries'] and recovery['source'] == build['source']
assert len(recovery['relocated_partial_files']) == 2
for row in recovery['relocated_partial_files']:
    path = artifact(row['retained'])
    assert path.parent == P / 'assembly-failed-0'
    assert path.name in ('control.scanner.symbol.txt', 'control.scanner.assembly.txt')
    assert row['original']['path'] == str(P / 'assembly' / path.name)
    assert all(row['original'][k] == row['retained'][k] for k in ('bytes', 'sha256'))
assert recovery['recorded_at'] >= spans[-1][1]

assembly = read(P / 'assembly/receipt.json')
assert assembly['schema'] == 'litchi.performance.0814.codegen.v1'
assert [(r['variant'], r['name']) for r in assembly['rows']] == [
    (variant, name) for variant in ('control', 'profile', 'fp') for name in ('scanner', 'inspector')]
assembly_end = spans[-1][1]
for row in assembly['rows']:
    assert row['binary'] == build['binaries'][row['variant']] and row['source'] == build['source']
    assert row['nm']['exit_code'] == row['objdump']['exit_code'] == 0
    assert row['nm']['matched_count'] == 1
    symbol = artifact(row['nm']['symbols']).read_text().split()
    assert len(symbol) == 4 and symbol[3] == row['symbol']
    assert int(symbol[0], 16) == row['address'] and int(symbol[1], 16) == row['size']
    assert row['size'] > 0
    text = artifact(row['objdump']['assembly']).read_text()
    instructions = [int(m[1], 16) for m in re.finditer(r'^\s*([0-9a-f]+):\s+(.+)$', text, re.M)]
    assert instructions and all(row['address'] <= n < row['address'] + row['size'] for n in instructions)
    assert spans[-1][1] <= row['nm']['started'] <= row['nm']['ended'] <= row['objdump']['started'] <= row['objdump']['ended']
    assembly_end = max(assembly_end, row['objdump']['ended'])

cleanup_path = P / 'cleanup.json'
if '--final' in sys.argv or cleanup_path.exists():
    cleanup = read(cleanup_path)
    assert cleanup['schema'] == 'litchi.performance.0814.cleanup.v1'
    assert cleanup['target'] == str(TARGET) and cleanup['target_removed'] is True
    assert cleanup['binaries_verified_before_removal'] is True and not TARGET.exists()
    assert cleanup['started'] >= assembly_end and cleanup['ended'] >= cleanup['started']
    assert len(cleanup['removed_binaries']) == 3
    assert sorted(cleanup['removed_binaries'], key=lambda r: r['path']) == sorted(build['binaries'].values(), key=lambda r: r['path'])
    assert artifact(cleanup['source_manifest']) == P / 'build/source.json'
else:
    for row in build['binaries'].values():
        artifact(row)
assert not list(P.rglob('__pycache__'))
print('0814 aggregate PASS:56 reports/1820 samples, exact independent timing/frame agreement, source and chronology')
