"""Complete frozen B1, B2, A2 serial order with shared-host observations."""
import subprocess
from run import HERE, capture, now, write

for stage, repeat in [('candidate', 1), ('candidate', 2), ('baseline', 2)]:
    rows = subprocess.check_output(['ps', '-eo', 'pid,ppid,etime,comm'], text=True).splitlines()
    rows = [r for r in rows if r.split()[-1] in ['cargo', 'rustc']]
    write(HERE/(f'native-{stage}-r{repeat}-process-observation.json'), dict(
        observed_utc=now(), processes=rows,
        scope='Pre-capture cargo/rustc snapshot, not proof of an idle host; external work uncontrolled.'))
    capture(stage, 'native', repeat)
