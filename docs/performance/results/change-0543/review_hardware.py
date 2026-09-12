"""Replay individually reviewed whole-child hardware diagnostic flags."""
import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent


def review():
    path = HERE / 'hardware-analysis.json'
    data = json.loads(path.read_text())
    flags = []
    for row in data['comparison']['rows']:
        for group in ('event_metrics', 'group_metrics'):
            for metric, values in row[group].items():
                change = values['change_percent']
                adverse = change < -5 if metric == 'ipc' else change > 5
                if adverse:
                    flags.append(dict(kind='candidate_adverse', shape=row['shape'],
                                      repeat=row['repeat'], metric=metric,
                                      baseline=values['baseline'], candidate=values['candidate'],
                                      change_percent=change))
    for stage, stage_data in data['stages'].items():
        captures = {(x['shape'], x['repeat']): x for x in stage_data['captures']}
        for shape in sorted({x[0] for x in captures}):
            first, second = captures[shape, 1], captures[shape, 2]
            for metric in sorted(first['events']) + ['ipc', 'branch_miss_percent']:
                if metric in first['events']:
                    a, b = first['events'][metric]['value'], second['events'][metric]['value']
                else:
                    a, b = first['group_metrics'][metric], second['group_metrics'][metric]
                change = (b / a - 1) * 100 if a else (0 if b == 0 else None)
                if change is None or abs(change) > 5:
                    flags.append(dict(kind='same_build_drift', stage=stage, shape=shape,
                                      metric=metric, first=a, second=b, change_percent=change))
    for i, flag in enumerate(flags):
        flag['id'] = f'0543-hardware-{i+1:03d}'
        metric = flag['metric']
        if metric in ('page-faults', 'context-switches', 'cpu-migrations'):
            meaning = 'Process/OS diagnostic; this single child counter does not establish an operation-local resource regression or its cause.'
        elif metric == 'ipc':
            meaning = 'Higher IPC is beneficial; changes describe the whole child and cannot establish isolated planning efficiency.'
        elif metric == 'branch_miss_percent':
            meaning = 'Higher branch miss percentage is adverse; the denominator can change with executed work, so this is not a planning-only branch claim.'
        else:
            meaning = 'Whole-child hardware count; setup, warmup and other workflow phases are included. Do not treat it as isolated planning cost.'
        flag['interpretation'] = meaning + (' Absolute repeat drift remains visible without an asserted noise cause.' if flag['kind'] == 'same_build_drift' else ' Retain this individual adverse comparison despite lower native workflow medians.')
        flag['disposition'] = 'reviewed; candidate rejected by cap latency and Clippy; no retained runtime claim'
    return dict(schema='litchi.0543.hardware-review.v1', status='complete',
                hardware_analysis_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                derivation_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
                threshold_percent=5, ipc_adverse_direction='decrease',
                other_metrics_adverse_direction='increase', same_build_drift_direction='absolute',
                scope='All seven measured counters plus IPC and branch miss percentage; whole-child diagnostics only.',
                adverse_count=sum(x['kind'] == 'candidate_adverse' for x in flags),
                drift_count=sum(x['kind'] == 'same_build_drift' for x in flags), flags=flags)


if __name__ == '__main__':
    result = review()
    target = HERE / 'hardware-review.json'
    if target.exists():
        assert json.loads(target.read_text()) == result
        print('hardware review exact replay passed')
    else:
        target.write_text(json.dumps(result, indent=2) + '\n')
        print(result['adverse_count'], result['drift_count'])
