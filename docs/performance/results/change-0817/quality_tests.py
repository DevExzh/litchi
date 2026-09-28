"""Recover only the failed feature-matrix test; reuse unchanged successful suites."""
import os
import re
import subprocess
import sys
import time
from pathlib import Path

import custody as c

assert len(sys.argv) == 2
out = Path(sys.argv[1]).resolve()
assert out.parent == c.P and out.name.startswith('quality-')
assert not (out / 'test-recovery.json').exists()
old = c.P / 'quality-1'
old_source = c.read(old / 'source.json')
current = {'production': c.source(), 'tool': c.tool_source()}
changed_test = 'tools/perf-baseline/tests/xlsx_planning_allocations.rs'
assert current['production'] == old_source['production']
assert set(current['tool']) == set(old_source['tool'])
changed = [n for n in current['tool'] if current['tool'][n] != old_source['tool'][n]]
assert changed == [changed_test]
original = old / 'input-snapshot/xlsx_planning_allocations.rs'
assert c.sha(original) == old_source['tool'][changed_test]
assert c.assert_root_inputs() == c.read(old / 'frozen-inputs.json')['root_inputs']
for name, digest in c.read(old / 'frozen-inputs.json')['packet'].items():
    assert c.sha(old / 'input-snapshot' / name) == digest, name
log = old / '02.log'
checks = c.read(old / 'checks.json')
assert checks[2]['log'] == c.artifact(log)
assert checks[2]['exit_code'] != 0
assert '--all-features' in checks[2]['command'] and '--test' not in checks[2]['command']
text = log.read_text()
suites = re.findall(r'test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out', text)
assert len(suites) == 27
assert sum(int(row[1]) for row in suites) == 640
assert sum(int(row[2]) for row in suites) == 1
assert sum(int(row[3]) for row in suites) == 1
assert all(row[0] == 'ok' and row[2] == '0' for row in suites[:-1])
assert suites[-1] == ('FAILED', '0', '1', '0', '0', '0')
suite_records = []
active_suite = None
for line in text.splitlines():
    match = re.match(r"\s*Running (.+) \((.+)\)$", line)
    if match:
        active_suite = match.group(1)
    match = re.match(r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out", line)
    if match:
        assert active_suite is not None
        suite_records.append({'suite': active_suite, 'status': match.group(1), 'passed': int(match.group(2)), 'failed': int(match.group(3)), 'ignored': int(match.group(4))})
assert len(suite_records) == 27
assert suite_records[-1]['suite'] == 'tests/xlsx_planning_allocations.rs'
assert all(row['status'] == 'ok' for row in suite_records[:-1])
assert 'to rerun pass `--test xlsx_planning_allocations`' in text
assert 'left: String("ordinary_save_procfs_operation_scoped")' in text
assert 'right: "none"' in text
assert sorted(x.name for x in (c.TOOL / 'tests').glob('*.rs')) == ['xlsx_filesystem.rs', 'xlsx_planning_allocations.rs']
manifest = str(c.TOOL / 'Cargo.toml')
base = ['cargo', 'test', '--offline', '--locked', '--manifest-path', manifest]
commands = [
    base + ['--all-features', '--test', 'xlsx_planning_allocations', '--', '--test-threads=2'],
    base + ['--features', 'allocator-metrics', '--test', 'xlsx_planning_allocations', '--', '--test-threads=2'],
    base + ['--all-features', '--doc', '--', '--test-threads=2'],
]
rows = []
for index, command in enumerate(commands):
    output = out / f'test-recovery-{index}.log'
    assert not output.exists()
    started = time.time()
    with output.open('w') as stream:
        result = subprocess.run(command, cwd=c.ROOT, stdout=stream, stderr=subprocess.STDOUT)
    rows.append({'command': command, 'started': started, 'ended': time.time(), 'exit_code': result.returncode, 'log': c.artifact(output), 'environment': {key: os.environ.get(key) for key in ('CARGO_TARGET_DIR', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_DEV_DEBUG', 'RUSTDOCFLAGS')}})
    c.write(out / 'test-recovery-commands.json', rows)
    assert result.returncode == 0, f'recovered test command failed: {output}'
    c.unchanged(current)
c.write(out / 'test-recovery.json', {
    'schema': 'litchi.performance.0817.test-recovery.v1',
    'inherited_passed': 640, 'inherited_ignored': 1,
    'inherited_successful_suites': 26, 'carried_suites': suite_records[:-1],
    'old_plan_sha256': c.sha(old / 'input-snapshot/plan.json'),
    'amended_plan_sha256': c.sha(c.P / 'plan.json'),
    'prior_log': c.artifact(log), 'prior_source': c.artifact(old / 'source.json'),
    'prior_frozen_inputs': c.artifact(old / 'frozen-inputs.json'),
    'changed_test': changed_test, 'old_test_sha256': old_source['tool'][changed_test],
    'new_test_sha256': current['tool'][changed_test],
    'runtime_and_production_unchanged': True, 'rows': rows,
})
print('0817 recovered tests PASS; 640 prior passes/1 ignored retained on unchanged suite sources')
