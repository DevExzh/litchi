"""Capture symbol-bounded CFB code from the actual measured executable."""
import argparse
import run as R


def inspect(stage):
    R.configure(stage)
    binary = R.SCRATCH / 'normal'
    owners = ('validate_physical_sector_layout', 'validate_stream_allocations')
    path = R.FOLDER / 'assembly-plan.json'
    assert not path.exists()
    R.write(path, dict(schema='litchi-0534-assembly-plan-v1',
        status='frozen-before-inspection', owners=list(owners),
        binary_sha256=R.sha(binary),
        source_manifest_sha256=R.sha(R.FOLDER / 'source-manifest.json'),
        script_sha256=R.sha(R.HERE / 'inspect_assembly.py'),
        scope='Static code shape and explicit caller edges; no static-instruction latency inference'))
    R.run('symbols', ['nm', '-S', '--defined-only', str(binary)], binary)
    rows = []
    for line in (R.FOLDER / 'symbols.stdout').read_text().splitlines():
        fields = line.split()
        if len(fields) != 4:
            continue
        address, size, kind, symbol = fields
        if kind not in ('t', 'T'):
            continue
        owner = next((o for o in sorted(owners, key=len, reverse=True) if o in symbol), None)
        if owner is None:
            continue
        name = 'assembly-' + owner + '-' + str(sum(r['owner'] == owner for r in rows))
        R.run(name, ['objdump', '-d', '--disassemble=' + symbol, str(binary)], binary)
        rows.append(dict(name=name, owner=owner, symbol=symbol, address_hex=address,
            size_bytes=int(size,16), receipt_sha256=R.sha(R.FOLDER / (name+'.receipt.json'))))
    assert {'validate_stream_allocations','validate_physical_sector_layout'} <= {r['owner'] for r in rows}
    R.write(R.FOLDER / 'assembly-index.json', dict(schema='litchi-0534-assembly-index-v1',
        plan_sha256=R.sha(path), symbols_receipt_sha256=R.sha(R.FOLDER / 'symbols.receipt.json'), rows=rows))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('stage', choices=['baseline','candidate'])
    inspect(parser.parse_args().stage)
