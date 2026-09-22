#!/usr/bin/env python3
"""Offline corruption checks against fresh exact-oracle qualification reports."""
import copy
from contract import P, read, validate_report, write

native = read(P/'qualification-restored/native-secondary-strict-drained.json')
allocation = read(P/'qualification-restored/allocation-secondary-strict-drained.json')
native_row = dict(case='secondary',build='restored',lifecycle='strict-drained',lane='native',samples=1,warmups=3)
allocation_row = dict(native_row,lane='allocation',warmups=0)
checks = [
    ('source-hash',native,native_row,lambda j:j.update(source_sha256='0'*64)),
    ('output-hash',native,native_row,lambda j:j['samples'][0].update(output_sha256='0'*64)),
    ('nested-oracle',native,native_row,lambda j:j['samples'][0]['oracle'].update(output_stream_bytes_match_expected=False)),
    ('oracle-control',native,native_row,lambda j:j['oracle_controls'][0].update(rejected=False)),
    ('missing-warmup',native,native_row,lambda j:j['warmup_receipts'].pop()),
    ('warmup-oracle',native,native_row,lambda j:j['warmup_receipts'][0]['oracle'].update(semantic_reopen_ok=False)),
    ('warmup-phase-missing',native,native_row,lambda j:j['warmup_receipts'][0].pop('phase_ns')),
    ('warmup-phase-null',native,native_row,lambda j:j['warmup_receipts'][0].update(phase_ns=None)),
    ('warmup-index-type',native,native_row,lambda j:j['warmup_receipts'][0].update(index=False)),
    ('sample-extra-field',native,native_row,lambda j:j['samples'][0].update(unqualified_field=1)),
    ('top-retention',native,native_row,lambda j:j.update(retained_witness_count=1)),
    ('sample-retention',native,native_row,lambda j:j['samples'][0].update(retained_witness_count=1)),
    ('lifecycle',native,native_row,lambda j:j.update(lifecycle='strict-retained')),
    ('sample-index',native,native_row,lambda j:j['samples'][0].update(index=1)),
    ('negative-duration',native,native_row,lambda j:j['samples'][0]['phase_ns'].update(whole_ns=-1)),
    ('allocation-timing-claim',allocation,allocation_row,lambda j:j.update(timing_claim=True)),
    ('negative-allocation',allocation,allocation_row,lambda j:j['samples'][0]['allocations']['whole'].update(allocated_bytes=-1)),
]
results = []
for name, reference, row, mutate in checks:
    validate_report(reference,row)
    changed = copy.deepcopy(reference)
    mutate(changed)
    try:
        validate_report(changed,row)
    except AssertionError as error:
        results.append(dict(name=name,status='rejected',message=str(error)))
    else:
        raise AssertionError('corruption accepted: '+name)
write(P/'negative-contract.json',dict(status='passed',checks=results))
print('PASS',len(results),'offline report corruption controls')
