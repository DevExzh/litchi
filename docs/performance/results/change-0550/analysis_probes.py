"""Reject corrupted copies of a real report without editing retained captures."""
import copy
import json
from pathlib import Path
import tempfile
import analyze_metrics as A
import run as R

plan = A.plan_data()
binary = A.binary_metadata('normal', plan)
job = next(A.CAP.jobs('native'))
source = R.FOLDER / (job['name'] + '.json')
catalog = R.FOLDER / (job['name'] + '.catalog.json')
original = json.loads(source.read_text())
rows = []
with tempfile.TemporaryDirectory(prefix='analysis-probes-', dir=R.TARGET / 'tmp') as folder:
    probe = Path(folder) / 'report.json'
    probe.write_text(json.dumps(original))
    A.validate_report(source, catalog, job, 'normal', binary, plan)
    for name in ['short_elapsed_vector', 'wrong_binary_hash', 'short_commit_vector']:
        document = copy.deepcopy(original)
        if name == 'short_elapsed_vector':
            document['results'][0]['elapsed_ns']['samples'].pop()
        elif name == 'wrong_binary_hash':
            document['binary_identity']['binary_sha256'] = '0' * 64
        else:
            document['results'][0]['source']['xlsx_cell_values']['commit_ns'].pop()
        probe.write_text(json.dumps(document))
        try:
            A.validate_report(probe, catalog, job, 'normal', binary, plan)
        except A.EvidenceError as error:
            rows.append({'case': name, 'rejected': True, 'error': str(error)})
        else:
            raise AssertionError(name + ' was accepted')
R.write(R.HERE / 'analysis-probes.json', {'status': 'pass',
    'script_sha256': R.sha(Path(__file__)), 'analyzer_sha256': R.sha(R.HERE / 'analyze_metrics.py'),
    'source_report_sha256': R.sha(source), 'source_report': str(source.relative_to(R.HERE)),
    'valid_control_passed': True, 'negative_cases': rows, 'temporary_copies_removed': True})
print('Valid control and three corrupted-report rejection probes passed.')
