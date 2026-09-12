"""Show that foreign and derived hints still enter full source XML validation.

These are debug integration-test guard profiles, not latency or speedup data.
The test executable is taken from the completed all-feature OOXML test log.
"""
import datetime
import json
import re
import subprocess
import time

from run import HERE, REPO, sha, sources, write
from analyze_profiles import parse_profile

FUNCTION = 'litchi_opc::source_backed::PartView::source_xml_with_hint'
CASES = ['equal_version_foreign_lineage_falls_back_to_current_original',
         'derived_hint_returns_current_original_and_keeps_derived_bytes_live']


def main():
    directory = HERE / 'fallback-profiles'
    directory.mkdir(exist_ok=True)
    manifest = json.loads((HERE / 'candidate/source-manifest.json').read_text())
    assert sources() == manifest
    log = (HERE / 'checks/ooxml-tests.log').read_text()
    paths = re.findall(r'Running tests/source_xml_hint.rs \(([^)]+)\)', log)
    assert len(paths) == 1
    from pathlib import Path
    binary = Path(paths[0])
    binary_hash = sha(binary)
    plan = dict(scope=__doc__, function=FUNCTION, cases=CASES,
                                       test_log_sha256=sha(HERE / 'checks/ooxml-tests.log'),
                                       binary=str(binary), binary_sha256=binary_hash)
    if not (directory / 'plan.json').exists():
        write(directory / 'plan.json', plan)
    else:
        assert json.loads((directory / 'plan.json').read_text()) == plan
    analyses = []
    for case in CASES:
        raw, stdout, stderr = [directory / (case + ext) for ext in ['.callgrind', '.stdout', '.stderr']]
        command = ['taskset', '-c', '2', 'valgrind', '--tool=callgrind', '--collect-atstart=no',
                   '--toggle-collect=*' + FUNCTION, '--callgrind-out-file=' + str(raw),
                   str(binary), '--exact', case, '--nocapture', '--test-threads=1']
        receipt_path = directory / (case + '.json')
        if receipt_path.exists():
            record = json.loads(receipt_path.read_text())
            assert record['command'] == command and record['exit_code'] == 0
            assert record['binary_sha256'] == binary_hash
            assert record['source_manifest_sha256'] == sha(HERE / 'candidate/source-manifest.json')
            assert all(sha(directory / name) == value for name, value in record['artifacts'].items())
        else:
            started = datetime.datetime.now(datetime.timezone.utc).isoformat()
            tick = time.monotonic()
            with stdout.open('x') as out, stderr.open('x') as err:
                result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
            record = dict(command=command, started_utc=started, elapsed_seconds=time.monotonic()-tick,
                          exit_code=result.returncode, source_unchanged=sources() == manifest,
                          source_manifest_sha256=sha(HERE / 'candidate/source-manifest.json'),
                          binary_sha256=sha(binary),
                          artifacts={p.name: sha(p) for p in [raw, stdout, stderr] if p.exists()})
            write(receipt_path, record)
        assert record['exit_code'] == 0 and record['source_unchanged']
        assert '1 passed; 0 failed' in stdout.read_text()
        # These focused tests create only in-memory fixtures. One exact test
        # invokes the measured method once, with no warmup or repeated sample.
        record.update(samples=1, warmups=0, repeats=1, cleanup_verified=True,
                      fixture_scope='In-memory integration-test fixtures; no scratch corpus files')
        receipt_path.write_text(json.dumps(record, indent=2)+'\n')
        profile = parse_profile(raw, FUNCTION)
        analyses.append(profile)
        print(case, 'scope passed', flush=True)
    write(directory / 'analysis.json', dict(scope=__doc__, profiles=analyses))


if __name__ == '__main__':
    main()
