"""Run source/tool gates after the serialized measurement phases."""
from pathlib import Path
import datetime
import json
import os
import subprocess

repo = Path('/home/zhuhe/code/litchi')
out = repo / 'docs/performance/results/change-0417/checks'
journal = out / 'gates.json'
records = json.loads(journal.read_text()) if journal.exists() else []
env = os.environ.copy()
env['RUSTUP_TOOLCHAIN'] = '1.98.1'
env['PYTHONDONTWRITEBYTECODE'] = '1'
commands = [
    ('python-tests', ['python3', '-m', 'unittest', 'tools.test_summarize_crud_baseline',
                     'tools.test_crud_coverage_index', 'tools.test_perf_compare',
                     'tools.test_perf_corpus_binding']),
    ('crud-index', ['python3', 'tools/validate_crud_coverage_index.py']),
    ('boundaries', ['python3', 'tools/check_crate_boundaries.py']),
    ('classification', ['python3', 'tools/check_report_claim_classification.py']),
]
for label, argv in commands:
    sequence = 1 + sum(record['label'] == label for record in records)
    log = out / f'{label}-{sequence}.log'
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with log.open('wb') as stream:
        process = subprocess.run(argv, cwd=repo, env=env, stdout=stream, stderr=subprocess.STDOUT)
    records.append(dict(label=label, sequence=sequence, argv=argv, cwd=str(repo),
                        started_utc=started,
                        finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                        exit_code=process.returncode, log=log.relative_to(out.parent).as_posix()))
    journal.write_text(json.dumps(records, indent=2) + '\n')
    print(label, sequence, process.returncode, flush=True)
    if process.returncode:
        raise SystemExit(process.returncode)
