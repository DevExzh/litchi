#!/usr/bin/env python3
"""Pin all corpus identities from matching final binary pilots before captures."""
from common import ROOT, read, write, meta, now


def allocation(value):
    assert value['status'] == 'measured'
    assert value['failed_allocation_calls'] == 0
    assert value['live_bytes_before'] + value['allocated_bytes'] - value['deallocated_bytes'] == value['live_bytes_after']
    assert value['region_peak_live_bytes'] >= max(value['live_bytes_before'], value['live_bytes_after'])


def sample_checks(value, instrumentation, mode):
    phases = ('open', 'snapshot', 'stage', 'commit', 'publish', 'drop')
    assert value['config']['samples'] == value['config']['warmups'] == 1
    assert value['config']['mode'] == mode
    for case in value['cases']:
        rows = case['total_samples' if mode == 'total' else 'phase_samples']
        assert len(rows) == 1 and rows[0]['sample'] == 0
        row = rows[0]
        assert row['sink']['accepted_bytes'] == case['corpus']['candidate_archive_bytes']
        assert row['sink']['sha256'] == case['corpus']['candidate_archive_sha256']
        assert row['sink']['largest_write'] <= 16384
        if mode == 'total':
            if instrumentation == 'allocator':
                allocation(row['allocation'])
                assert row['allocation']['live_bytes_before'] == row['allocation']['live_bytes_after']
            else:
                assert row['allocation'] is None
        else:
            for field in ('calls', 'requested_bytes', 'returned_bytes'):
                assert sum(row[phase]['source_reads'][field] for phase in phases) == row['source_reads'][field]
            for field, expected in row['source_reads']['request_histogram'].items():
                assert sum(row[phase]['source_reads']['request_histogram'][field] for phase in phases) == expected
            if instrumentation == 'allocator':
                for phase in phases:
                    allocation(row[phase]['allocation'])
                for before, after in zip(phases, phases[1:]):
                    assert row[before]['allocation']['live_bytes_after'] == row[after]['allocation']['live_bytes_before']
                assert row['open']['allocation']['live_bytes_before'] == row['drop']['allocation']['live_bytes_after']
            else:
                assert all(row[phase]['allocation'] is None for phase in phases)


def main():
    cases = None
    pilots = {}
    for instrumentation in ('normal', 'allocator'):
        for mode in ('total', 'phases'):
            label = f'pilot-{instrumentation}-{mode}'
            report = ROOT / f'{label}.report.json'
            gate = read(ROOT / 'validation' / f'{label}.json')
            assert gate['exit_code'] == 0 and gate['source_unchanged']
            value = read(report)
            sample_checks(value, instrumentation, mode)
            current = {str(case['count']): case['corpus'] for case in value['cases']}
            assert set(current) == {'64', '8192', '131072'}
            if cases is None:
                cases = current
            assert current == cases
            pilots[label] = dict(path=report.name, **meta(report))
    write(ROOT / 'corpus-manifest.json', dict(
        schema='docx-plain-paragraph-tail-append-corpus-manifest-v1',
        frozen_utc=now(), cases=cases, pilots=pilots))
    print('All three corpus identities agree across both builds and both modes.')


if __name__ == '__main__':
    main()
