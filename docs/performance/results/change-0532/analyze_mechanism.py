"""Connect constructor instructions to corpus geometry and recorded assembly."""
import argparse
import json
from pathlib import Path

from run import HERE, sha


def read(name):
    return json.loads((HERE / name).read_text())


def analyze():
    profiles = read('profile-analysis.json')
    native = read('analysis.json')
    assembly = read('assembly-index.json')
    assert profiles['status'] == native['status'] == 'pass'
    code = []
    for row in assembly['rows']:
        if row['owner'] != 'claim_sector':
            continue
        path = HERE / 'baseline' / (row['name'] + '.stdout')
        instructions = [line.split('\t')[2].split('#')[0].strip()
                        for line in path.read_text().splitlines()
                        if len(line.split('\t')) >= 3 and line.split('\t')[2].strip()]
        entry, success = instructions[:10], instructions[-5:]
        assert [s.split()[0] for s in entry] == [
            'push', 'sub', 'mov', 'mov', 'mov', 'cmp', 'jbe', 'movzbl', 'test', 'je']
        assert [s.split()[0] for s in success] == ['mov', 'movq', 'add', 'pop', 'ret']
        assert '$0x70,%rsp' in entry[1] and '$0x70,%rsp' in success[2]
        assert '+0x15a>' in entry[-1] and row['size_bytes'] == 363
        code.append(dict(name=row['name'], symbol=row['symbol'], bytes=363,
                         stack_reservation_bytes=112,
                         success_path_instruction_count=len(entry) + len(success),
                         entry_instructions=entry, success_instructions=success,
                         assembly_sha256=sha(path)))
    assert len(code) == 6
    caller_rows = []
    for row in assembly['rows']:
        if row['owner'] != 'validate_stream_allocations':
            continue
        path = HERE / 'baseline' / (row['name'] + '.stdout')
        calls = [line.strip() for line in path.read_text().splitlines()
                 if '\tcall' in line and 'claim_sector' in line]
        caller_rows.append(dict(name=row['name'], claim_call_sites=calls,
                                assembly_sha256=sha(path)))
    assert any(r['claim_call_sites'] for r in caller_rows)
    rows = []
    for profile in profiles['profiles']:
        raw = read(profile['profile_result'])
        assert len(raw['results']) == 1
        corpus = raw['results'][0]['corpus']
        # Both generators use OleWriter::new(), whose source-bound default
        # is 512-byte sectors. This equation applies to these valid fixtures.
        assert corpus['archive_bytes'] % 512 == 0
        physical_sectors = corpus['archive_bytes'] // 512 - 1
        claims = []
        for dump in profile['constructor_attribution']:
            claim = dump['functions']['claim_sector']
            assert claim['calls'] == physical_sectors
            assert claim['self_ir'] == claim['inclusive_ir'] == 15 * physical_sectors
            assert claim['direct_ir'] == 0
            assert claim['validation']['positive_incoming_edges_retained']
            claims.append(dict(dump=dump['dump'], calls=claim['calls'],
                               exclusive_ir=claim['self_ir'], ir_per_call=15))
        assert len(claims) == 5
        ranking = profile['exclusive_sector_ranking']
        case = raw['results'][0]['case']
        allocations = [r for r in native['rows'] if r['lane'] == 'alloc'
                       and r['case'] == case and r['shape'] == corpus['shape']
                       and r['repeat'] == profile['repeat']]
        assert len(allocations) == 1
        allocation = allocations[0]['allocation']
        rows.append(dict(name=profile['name'], case=case, shape=corpus['shape'],
                         repeat=profile['repeat'], archive_bytes=corpus['archive_bytes'],
                         physical_sectors=physical_sectors, claims=claims,
                         constructor_ir=ranking['constructor_inclusive_ir'],
                         exclusive_owner_ranking=[{k: r[k] for k in (
                             'sector', 'exclusive_ir', 'share_percent', 'rank')}
                             for r in ranking['rows']],
                         allocation_unique_values={k: sorted(set(allocation[k])) for k in (
                             'allocation_calls', 'reallocation_calls', 'allocated_bytes',
                             'incremental_region_peak_live_bytes')}))
    return dict(schema='litchi-0532-claim-mechanism-v1', status='pass',
                performance_claim='none',
                bindings={n: sha(HERE / n) for n in (
                    'profile-analysis.json', 'analysis.json', 'assembly-index.json',
                    'baseline/source-manifest.json')},
                source_geometry={'generator': 'tools/perf-baseline/src/lib.rs',
                                 'writer_default': 'crates/litchi-cfb/src/writer/core.rs',
                                 'sector_bytes': 512},
                code=code, caller_code=caller_rows, rows=rows,
                next_candidate='Private cold error helpers and ordinary inlining for claim_sector; preserve exact checked body and collect-then-claim order.',
                limitations='15 Ir per claim is measured only in these constructor dumps and matches the recorded x86-64 success path. It is not a latency saving estimate, allocation count, or universal instruction cost. Inclusive parent costs overlap. The standalone harness has no explicit LTO profile; root workspace release LTO does not apply to it. No candidate is adopted.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', nargs='?', type=Path,
                        default=HERE / 'mechanism-analysis.json')
    args = parser.parse_args()
    value = analyze()
    args.output.write_text(json.dumps(value, indent=2) + '\n')
    print('Claim mechanism equations and assembly checks passed')
