"""Bind physical reconciliation and accounting symbols to measured binaries."""
import json
import sys
import run as R


def inspect(stage):
    R.configure(stage)
    binary = R.SCRATCH / 'normal'
    R.run('symbols', ['nm', '-S', '--defined-only', str(binary)], binary)
    rows = []
    owners = ('validate_physical_sector_layout', 'load_fat', 'claim_sector',
              'claim_chain', 'validate_stream_allocations', 'collect_exact',
              'open_with_limits', 'pending', 'advance')
    for line in (R.FOLDER / 'symbols.stdout').read_text().splitlines():
        fields = line.split()
        if len(fields) != 4:
            continue
        address, size, kind, symbol = fields
        if kind not in ('t', 'T') or 'litchi_cfb' not in symbol:
            continue
        if not any(owner in symbol for owner in owners):
            continue
        name = 'assembly-' + str(len(rows))
        R.run(name, ['objdump', '-d', '--disassemble=' + symbol, str(binary)], binary)
        rows.append(dict(name=name, symbol=symbol, address_hex=address,
                         size_bytes=int(size, 16),
                         receipt_sha256=R.sha(R.FOLDER / (name + '.receipt.json'))))
    assert rows
    R.write(R.FOLDER / 'assembly-index.json', dict(
        schema='ole2_0555_assembly_v1', plan_sha256=R.sha(R.HERE / 'plan.json'),
        binary_sha256=R.sha(binary),
        source_manifest_sha256=R.sha(R.FOLDER / 'source-manifest.json'),
        script_sha256=R.sha(R.HERE / 'inspect_assembly.py'), rows=rows,
        requested_owners=list(owners),
        scope='Static instructions of emitted matching symbols; absent helpers may be inlined. Dynamic Callgrind edges and positions remain required for work attribution.'))


if __name__ == '__main__':
    assert sys.argv[1] in ('baseline', 'candidate')
    inspect(sys.argv[1])
