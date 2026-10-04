"""Compile the exact help function and registry test without workspace dependencies."""
import hashlib
import json
from pathlib import Path
import re
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
source = (ROOT / 'tools/perf-baseline/src/lib.rs').read_text()
base = subprocess.check_output(['git', 'show', '2b974e3541517ccf032021dc5497243b315fd313:tools/perf-baseline/src/lib.rs'], cwd=ROOT, text=True)

def usage(text):
    return 'fn usage_text()' + text.split('fn usage_text()', 1)[1].split('\nfn print_usage', 1)[0]

def run(argv, **kwargs):
    result = subprocess.run(argv, capture_output=True, text=True, **kwargs)
    return dict(argv=argv, exit_code=result.returncode, stdout=result.stdout, stderr=result.stderr)

constants = '\n'.join(re.findall(r'^const DEFAULT_(?:SAMPLES|WARMUP_ITERATIONS|RANGE_[A-Z_]+):[^\n]+', source, re.M))
test = '    #[test]\n    fn selectable_case_registry_has_complete_name_and_parse_mappings()' + source.split('    fn selectable_case_registry_has_complete_name_and_parse_mappings()', 1)[1].split('\n    #[test]', 1)[0]
receipt = {'scope': 'Exact extracted help function and registry test; not the full Cargo harness suite.', 'source_sha256': hashlib.sha256(source.encode()).hexdigest(), 'toolchain': run(['rustc', '+1.95.0', '--version']), 'legs': []}
with tempfile.TemporaryDirectory(prefix='litchi-0906-help-') as tmp:
    path = Path(tmp)
    (path / 'lib.rs').write_text(source)
    for label, text in [('before_help_fix', base), ('after_help_fix', source)]:
        driver = constants + '\n' + usage(text) + '\n' + test
        (path / 'driver.rs').write_text(driver)
        compilation = run(['rustc', '+1.95.0', '--edition=2024', '--test', '--crate-name', 'help_registry', str(path / 'driver.rs'), '-o', str(path / 'test')])
        assert compilation['exit_code'] == 0, compilation
        execution = run([str(path / 'test'), '--nocapture'])
        assert (execution['exit_code'] == 0) == (label == 'after_help_fix'), execution
        receipt['legs'].append({'label': label, 'driver_sha256': hashlib.sha256(driver.encode()).hexdigest(), 'compile': compilation, 'test': execution})
receipt.update(status='pass', temporary_directory_removed=not path.exists(), worktrees_created=0)
with (PACKET / 'help-regression-verification.json').open('x') as out:
    json.dump(receipt, out, indent=2)
    out.write('\n')
print('PASS: old help rejected; corrected help accepted; temporary build removed')
