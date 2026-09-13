#!/usr/bin/env python3
"""Supplemental descriptive repeat review; never changes the frozen noise gate."""
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('analysis0557', HERE / 'analyze.py')
a = importlib.util.module_from_spec(spec)
spec.loader.exec_module(a)
plan = a.plan_data()
expected, shapes = a.expected_corpus()
rows = a.noise_rows(plan, expected, shapes)
groups = {(r['case'], r['shape'], r['repeat']): r for r in rows}
records = []
for case in a.PRIMARY:
    for shape in a.SHAPES:
        first, second = (groups[(case, shape, repeat)] for repeat in (1, 2))
        key = (case, shape)
        records.extend(a.timing_records(first, second, 'baseline_noise_repeat', key))
        for phase, values in first['source']['phases'].items():
            left, right = a.statistics(values), a.statistics(second['source']['phases'][phase])
            for statistic in a.COMPARE_STATS:
                def get(value):
                    if statistic.startswith('confidence_interval_95.'):
                        return value['confidence_interval_95'][statistic.split('.')[-1]]
                    return value[statistic]
                records.append(a.change_record(get(left), get(right), family='baseline_noise_repeat',
                               key=key, metric=f'phase.{phase}.{statistic}'))
        for field in a.RSS_FIELDS[1:]:
            records.append(a.change_record(first['rss'][field], second['rss'][field],
                           family='baseline_noise_repeat', key=key, metric=f'rss.{field}'))
flags = [r for r in records if r['over_five_percent']]
for r in flags:
    if r['metric'] == 'elapsed_ns.p50':
        r['review'] = 'This same-binary p50 drift participates in the preregistered maximum absolute drift N; the pilot is stopped, with no candidate measurement or noise rescue.'
    elif r['metric'].startswith('phase.'):
        r['review'] = 'Retained same-binary phase variability. The phase boundaries are unchanged across repeats; the cause is not established by this pilot and this row cannot attribute a candidate effect or rescue the gate.'
    elif r['metric'].startswith('rss.'):
        r['review'] = 'Retained whole-child diagnostic variability, including work outside the timed phases. It is not an operation allocation measurement and cannot attribute a candidate effect or rescue the gate.'
    else:
        r['review'] = 'Retained same-binary distribution variability. Means, tails, extrema, dispersion, and confidence endpoints do not override the preregistered p50 stop or establish a candidate effect.'
report = {
    'schema': 'xlsx_0557_noise_diagnostic_review_v1',
    'scope': 'Supplement added after the pilot to enumerate every >5% repeat drift in native elapsed and phase distribution statistics and all process sidecars. No gate or capture is changed.',
    'noise_analysis_sha256': hashlib.sha256((HERE/'noise-analysis.json').read_bytes()).hexdigest(),
    'analyzer_sha256': hashlib.sha256((HERE/'analyze.py').read_bytes()).hexdigest(),
    'review_script_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
    'comparison_count': len(records), 'flag_count': len(flags),
    'comparisons': records, 'reviewed_flags': flags,
    'allocation_scope': 'All five native allocation phase vectors are explicitly unavailable. Allocator integration passed separately; no allocator performance matrix or candidate capture ran.',
}
a.write_identical(HERE/'noise-review.json', report)
lines = ['# 0557 baseline repeat diagnostics', '', report['scope'], '',
         f"{len(records)} comparisons; {len(flags)} individual absolute repeat drifts above 5%.", '',
         'All are the same baseline binary. The cause of the drift is unestablished; no row is a candidate effect. The frozen maximum-p50 gate remains authoritative.', '',
         '| Case | Shape | Metric | R1 | R2 | Change | Review |',
         '| --- | --- | --- | ---: | ---: | ---: | --- |']
for r in flags:
    case, shape = r['key']
    delta = 'zero denominator' if r['change_percent'] is None else f"{r['change_percent']:+.3f}%"
    lines.append(f"| {case} | {shape} | {r['metric']} | {r['first']} | {r['second']} | {delta} | {r['review']} |")
text = '\n'.join(lines)+'\n'
path = HERE/'noise-review.md'
if path.exists():
    assert path.read_text() == text
else:
    with path.open('x') as f: f.write(text)
print(f'Noise diagnostics: {len(records)} comparisons; {len(flags)} individually retained/reviewed flags.')
