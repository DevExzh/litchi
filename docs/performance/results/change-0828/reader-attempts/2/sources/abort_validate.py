"""Offline validation of the pre-workload symbol-gate abort and static diagnosis."""
import json
from pathlib import Path
import re
import sys

import analysis
import custody as c
import validate


def main():
    assert sys.argv[1:] in ([], ['--final'])
    final = sys.argv[1:] == ['--final']
    cv = analysis.load_custody()
    quality = analysis.load_quality(cv)
    build = analysis.load_build(cv, quality)
    c.stable(build['frozen'])
    validate.check_reader_attempts()
    abort = c.read(c.P / 'abort.json')
    assert abort['schema'] == 'litchi.performance.0828.abort.v1'
    assert abort['status'] == 'aborted-before-workload'
    assert abort['reports'] == abort['samples'] == 0
    assert abort['failed_command'] == [
        'python3', '-B', 'docs/performance/results/change-0828/decode.py', '--symbols']
    assert abort['exit_code'] == 1
    for key in ('console', 'quality', 'build', 'frozen_inputs', 'failed_driver'):
        c.verify_descriptor(abort[key])
    assert abort['failed_driver']['sha256'] == build['frozen']['drivers']['decode.py']
    console = Path(abort['console']['path']).read_text()
    assert 'assert matched["symbol"] in assembly_text' in console
    assert console.rstrip().endswith('AssertionError')
    assert not any((c.P / name).exists() for name in (
        'qualification', 'native', 'perf', 'analysis.json', 'native.csv',
        'root-audit.json', 'frame-audit.json', 'stack-diagnostics.json',
        'symbols/symbol.json', 'symbols/complete.json'))
    for item in abort['partial_symbol_artifacts']:
        c.verify_descriptor(item)
    assert {str(path.relative_to(c.P)) for path in (c.P / 'symbols').iterdir()} == {
        str(Path(item['path']).relative_to(c.P)) for item in abort['partial_symbol_artifacts']}
    failed_assembly = (c.P / 'symbols/edit_region-assembly.txt').read_text()
    assert not re.search(r'^\s*[0-9a-fA-F]+:\s', failed_assembly, re.M)

    diagnostic_path = c.verify_descriptor(abort['diagnostic'])
    diagnostic = c.read(diagnostic_path)
    assert diagnostic['schema'] == 'litchi.performance.0828.symbol-diagnostic.v1'
    assert diagnostic['status'] == 'pass'
    assert diagnostic['workload_executed'] is diagnostic['capture_authorized'] is False
    assert diagnostic['binary'] == build['raw']['binaries']['fp']['artifact']
    c.verify_descriptor(diagnostic['driver'])
    c.verify_descriptor(diagnostic['build'])
    receipts = c.read(c.verify_descriptor(diagnostic['receipts']))
    assert len(receipts) == 6
    plan = cv['plan']
    owners = {'edit_region': plan['perf']['owner'], **plan['perf']['phase_owners']}
    assert set(diagnostic['proofs']) == set(owners)
    for receipt in receipts:
        assert receipt['exit_code'] == 0 and receipt['started'] <= receipt['ended']
        c.verify_descriptor(receipt['stdout'])
        assert not c.verify_descriptor(receipt['stderr']).read_bytes()
    binary_path = diagnostic['binary']['path']
    assert receipts[0]['command'] == ['nm', '-S', '--defined-only', binary_path]
    assert receipts[1]['command'] == ['nm', '-C', '-S', '--defined-only', binary_path]
    for receipt, (label, owner) in zip(receipts[2:], owners.items(), strict=True):
        proof = diagnostic['proofs'][label]
        assert receipt['label'] == label and proof['owner'] == owner
        raw, demangled = proof['raw'], proof['demangled']
        assert demangled['symbol'] == owner
        assert all(raw[key] == demangled[key] for key in ('address', 'size', 'type'))
        start, end = raw['address'], raw['address'] + raw['size']
        assert raw['size'] > 0
        assert receipt['command'] == [
            'objdump', '--demangle=rust', '--wide', '--line-numbers', '-d',
            f'--start-address={start}', f'--stop-address={end}', binary_path]
        assembly = c.verify_descriptor(proof['assembly']).read_text()
        assert proof['assembly'] == receipt['stdout']
        assert f'<{owner}>:' in assembly
        addresses = [int(match[1], 16) for line in assembly.splitlines()
                     if (match := re.match(r'^\s*([0-9a-fA-F]+):\s', line))]
        assert addresses and addresses[0] == start
        assert len(addresses) == proof['instruction_count']
        assert all(start <= address < end for address in addresses)
        assert re.search(r'push\s+%rbp', assembly)
        assert re.search(r'mov\s+%rsp,%rbp', assembly)
        assert re.search(r'\bcall\b', assembly)
    if final:
        cleanup = c.read(c.P / 'cleanup.json')
        validate.validate_cleanup(cleanup, build)
        assert (c.P / 'results-review.md').is_file()
        assert (c.P / 'symbol-failure-review.md').is_file()
    print(json.dumps({'status': abort['status'], 'reports': 0, 'samples': 0,
                      'fresh_probe_tests': 3, 'static_wrappers_verified': 4,
                      'final': final}, sort_keys=True))


if __name__ == '__main__':
    main()
