#!/usr/bin/env python3
"""Exercise both frozen validators on synthetic copies of qualified reports.

Synthetic data never enters the real capture directory. Source/binary custody
continues to use the real frozen root; schema/statistics use a temporary packet.
"""
import copy
import importlib.util
import json
from pathlib import Path
import shutil
import tempfile
from contract import P, command, read, sha, write
from run import guard


def load(name):
    spec = importlib.util.spec_from_file_location('preflight_' + name, P / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


guard()
assert not (P / 'preflight.json').exists()
with tempfile.TemporaryDirectory(prefix='litchi-0738-preflight-') as directory:
    shadow = Path(directory)
    for name in ('plan.json', 'cases.json', 'freeze.json'):
        shutil.copy2(P/name, shadow/name)
    write(shadow/'preflight.json',dict(status='passed',synthetic=True,freeze_sha256=sha(P/'freeze.json')))
    captures = shadow/'captures'
    captures.mkdir()
    rows = []
    for position, row in enumerate(read(P/'plan.json')['schedule']):
        suffix = '-w3' if row['lane']=='allocation' and row['warmups']==3 else ''
        report = read(P / ('qualification-'+row['build']) /
                      f'{row["lane"]}-{row["case"]}-{row["lifecycle"]}{suffix}.json')
        report['samples_requested'] = row['samples']
        sample = report['samples'][0]
        report['samples'] = []
        retained = row['lifecycle'] != 'strict-drained'
        if row['build']!='archive':
            report['retained_witness_count'] = row['samples'] if retained else 0
        for index in range(row['samples']):
            item = copy.deepcopy(sample)
            item['index'] = index
            if row['build']!='archive':
                item['retained_witness_count'] = index+1 if retained else 0
            if row['lane']=='native':
                item['phase_ns']['whole_ns'] = 1_000_000 + position*1000 + index*100
            report['samples'].append(item)
        name = f'{position:03d}.json'
        (captures/name).write_text(json.dumps(report,separators=(',',':'))+'\n')
        stderr = (captures/name).with_suffix('.stderr')
        stderr.write_bytes(b'')
        rows.append(dict(**row,command=command(row),exit_code=0,start_monotonic_ns=position*2+1,
                         end_monotonic_ns=position*2+2,output=name,sha256=sha(captures/name),
                         stderr_sha256=sha(stderr)))
    write(captures/'manifest.json',dict(status='complete',freeze_sha256=sha(P/'freeze.json'),
          preflight_sha256=sha(shadow/'preflight.json'),runs=rows))
    analyzer = load('analyze')
    audit = load('audit')
    analyzer.P = shadow
    audit.P = shadow
    audit.CAPTURES = captures
    analyzer.main()
    audit.main()
    original = read(shadow/'analysis.json')
    changed = copy.deepcopy(original)
    changed['comparisons'][0]['metrics']['p50']['median'] += 1
    write(shadow/'analysis.json',changed)
    try:
        audit.main()
    except audit.AuditError:
        corrupted_statistic_rejected = True
    else:
        raise AssertionError('independent audit accepted altered median')
write(P/'preflight.json',dict(status='passed',synthetic_processes=96,synthetic_native_samples=3600,
      freeze_sha256=sha(P/'freeze.json'),analyzer_sha256=sha(P/'analyze.py'),
      audit_sha256=sha(P/'audit.py'),corrupted_statistic_rejected=corrupted_statistic_rejected,
      temporary_packet_removed=True,scope='Schema and statistical integration only; no native timing evidence.'))
print('PASS synthetic 96-process preflight and independent altered-statistic rejection')
