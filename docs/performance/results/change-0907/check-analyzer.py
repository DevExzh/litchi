"""Reject damaged sample and guard evidence using an otherwise real raw report."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
import tempfile

sys.dont_write_bytecode = True
P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('pptx_analysis', P / 'analyze.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
raw = json.loads((P / 'eager-a.json').read_text())
checks = []
with tempfile.TemporaryDirectory(prefix='litchi-0907-analyzer-') as tmp:
    directory = Path(tmp)
    def check(label, mutate, expected):
        value = copy.deepcopy(raw)
        mutate(value)
        (directory / 'eager-a.json').write_text(json.dumps(value))
        try:
            module.raw_report(directory, 'eager-a')
        except module.EvidenceError as error:
            assert expected and expected in str(error), (label, str(error))
            checks.append({'case': label, 'rejected': True, 'reason': str(error)})
        else:
            assert expected is None, label
            checks.append({'case': label, 'accepted': True})
    check('unmodified', lambda r: None, None)
    check('missing_sample', lambda r: r['results'][0]['elapsed_ns']['samples'].pop(), '200')
    check('bad_order', lambda r: r['results'][0]['elapsed_ns']['sample_order'].__setitem__(0, r['results'][0]['elapsed_ns']['sample_order'][1]), 'permutation')
    check('wrong_p50', lambda r: r['results'][0]['elapsed_ns'].__setitem__('p50', r['results'][0]['elapsed_ns']['p50'] + 1), 'p50')
    check('wrong_materializations', lambda r: r['results'][0]['source']['ordinary_payload_materializations'].__setitem__(0, 228), 'materialization')
    check('zero_payload_reads', lambda r: r['results'][0]['source']['ordinary_payload_read_calls'].__setitem__(0, 0), 'no source read')
    check('oversized_sink_write', lambda r: r['results'][0]['sink'].__setitem__('largest_write', 65537), 'largest write')
    check('missing_binary_identity', lambda r: r.pop('binary_identity'), 'binary_identity')
with (P / 'analyzer-checks.json').open('x') as out:
    json.dump({'status': 'pass', 'helper_sha256': hashlib.sha256((P / 'analyze.py').read_bytes()).hexdigest(), 'raw_report_sha256': hashlib.sha256((P / 'eager-a.json').read_bytes()).hexdigest(), 'checks': checks, 'temporary_directory_removed': not directory.exists(), 'scope': 'Offline raw-report semantic and sample checks; no benchmark, artifact mutation, or build.'}, out, indent=2)
    out.write('\n')
print('analyzer checks pass')
