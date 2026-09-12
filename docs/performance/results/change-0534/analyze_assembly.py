"""Replay symbol/caller code evidence from the actual measured stage binaries."""
import argparse
import json
from pathlib import Path
import re
import analyze as A


def stage_report(stage):
    A.configure(stage)
    folder = A.FOLDER
    binary = A.binary('normal')
    plan_path = folder / 'assembly-plan.json'
    plan = A.read(plan_path)
    assert plan['script_sha256'] == A.sha(A.HERE / 'inspect_assembly.py')
    assert plan['binary_sha256'] == binary['sha256']
    assert plan['source_manifest_sha256'] == A.sha(folder / 'source-manifest.json')
    index = A.read(folder / 'assembly-index.json')
    assert index['plan_sha256'] == A.sha(plan_path)
    assert index['symbols_receipt_sha256'] == A.sha(folder / 'symbols.receipt.json')
    symbols_receipt = A.receipt('symbols', 'normal')
    assert symbols_receipt['command'] == ['nm', '-S', '--defined-only', binary['path']]
    parsed = {}
    for line in (folder / 'symbols.stdout').read_text().splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[2] in ('t', 'T'):
            address, size, _, symbol = parts
            owner = next((o for o in sorted(plan['owners'], key=len, reverse=True) if o in symbol), None)
            if owner:
                parsed[symbol] = (address, int(size, 16), owner)
    assert set(parsed) == {row['symbol'] for row in index['rows']}
    assert plan['owners'] == ['validate_physical_sector_layout', 'validate_stream_allocations']
    rows = []
    for row in index['rows']:
        assert parsed[row['symbol']] == (row['address_hex'], row['size_bytes'], row['owner'])
        receipt = A.receipt(row['name'], 'normal')
        assert row['receipt_sha256'] == A.sha(folder / (row['name'] + '.receipt.json'))
        assert receipt['command'] == ['objdump', '-d', '--disassemble=' + row['symbol'], binary['path']]
        raw = (folder / (row['name'] + '.stdout')).read_text()
        assert '<' + row['symbol'] + '>:' in raw
        rows.append(dict(**row,
            stack_reservations_hex=re.findall(r'\bsub\s+\$(0x[0-9a-f]+),%rsp', raw)))
    owners = {}
    for owner in plan['owners']:
        selected = [row for row in rows if row['owner'] == owner]
        owners[owner] = dict(symbol_count=len(selected),
            size_bytes=sum(row['size_bytes'] for row in selected),
            variant_sizes=[row['size_bytes'] for row in selected])
    return dict(binary_sha256=binary['sha256'], binary_bytes=binary['bytes'],
        source_manifest_sha256=A.sha(folder / 'source-manifest.json'),
        assembly_index_sha256=A.sha(folder / 'assembly-index.json'), owners=owners, rows=rows)


def analyze():
    return dict(status='pass', stages={stage: stage_report(stage) for stage in ('baseline','candidate')},
        scope='Symbol-bounded x86-64 static code sizes and stack reservations. Loop shape requires inspection of the bound disassembly; native and parent-constructor evidence are separate. No dynamic call count, removable instruction estimate or latency claim from static code.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('output', nargs='?', type=Path)
    args = parser.parse_args()
    output = args.output or A.HERE / 'assembly-analysis.json'
    output.write_text(json.dumps(analyze(), indent=2) + '\n')
    print('Matched assembly evidence verified')
