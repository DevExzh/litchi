"""Run all six repair gates serially after the focused regression."""
import re
import subprocess
import sys
from pathlib import Path

P = Path(__file__).resolve().parent
sys.path.insert(0, str(P.parent))
import custody as c

assert not (P / 'quality.json').exists()
focused = P / 'commands/focused/result.json'
assert c.read(focused)['exit_code'] == 0
names = ['fmt', 'check', 'tests', 'clippy', 'rustdoc', 'boundaries']
for name in names:
    subprocess.run([sys.executable, '-B', str(P / 'run.py'), name], cwd=c.ROOT, check=True)
rows = [c.artifact(P / 'commands' / name / 'result.json') for name in names]
c.write(P / 'quality.json', {'schema': 'litchi.performance.0820.repair-quality.v1',
                            'status': 'pass', 'gate_count': 6, 'gates': rows,
                            'focused': c.artifact(focused), 'source': c.artifact(P / 'source.json'),
                            'inputs': c.artifact(P / 'inputs.json')})
log = (P / 'commands/tests/output.log').read_text()
counts = re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out;', log)
assert counts and 'test result: FAILED.' not in log
summary = {'schema': 'litchi.performance.0820.repair-tests.v1', 'suites': len(counts),
           'passed': sum(int(r[0]) for r in counts), 'failed': sum(int(r[1]) for r in counts),
           'ignored': sum(int(r[2]) for r in counts), 'log': c.artifact(P / 'commands/tests/output.log'),
           'scope': 'full perf-baseline all-features tests including doctest invocation'}
c.write(P / 'test-summary.json', summary)
print('0820 repair quality complete', flush=True)
