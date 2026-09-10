#!/usr/bin/env python3
"""Compare historical scalar controls and explicit managed batch alternatives."""
import csv, hashlib, json, math, re
from collections import Counter
from pathlib import Path
HERE = Path(__file__).resolve().parent

def load(phase, name):
    path = HERE / phase / name
    receipt = json.loads(path.with_suffix('.json').read_text())
    assert receipt['exit_code'] == 0 and receipt['cleanup_verified']
    for suffix, key in [('.csv', 'csv_sha256'), ('.time.stderr', 'stderr_sha256')]:
        assert hashlib.sha256(path.with_suffix(suffix).read_bytes()).hexdigest() == receipt[key]
    rows = list(csv.DictReader(path.with_suffix('.csv').open()))
    assert len(rows) == 66
    assert {(int(r['repeat']), r['warmup'], int(r['ordinal'])) for r in rows} == {
        (repeat, warmup, i) for repeat in [0, 1] for warmup, count in [('true', 3), ('false', 30)] for i in range(count)}
    identity = {(r['fixture_sha256'], r['output_sha256'], r['output_bytes']) for r in rows}
    assert len(identity) == 1
    assert all(int(r['budget_after_memory']) == int(r['budget_after_objects']) == 0 for r in rows)
    rss = int(re.search(r'Maximum resident set size \(kbytes\):\s*(\d+)', path.with_suffix('.time.stderr').read_text()).group(1))
    return [r for r in rows if r['warmup'] == 'false'], rss, identity

def stats(rows):
    result = {}
    for field in ['elapsed_ns', 'edit_ns', 'open_ns', 'commit_ns', 'publish_ns', 'drop_ns']:
        values = sorted(int(r[field]) for r in rows)
        prefix = field.removesuffix('_ns')
        result.update({f'{prefix}_p{p}_us': values[math.ceil(len(values) * p / 100) - 1] / 1000 for p in [50, 95, 99]})
        result[f'{prefix}_mean_us'] = sum(values) / len(values) / 1000
    result['throughput_output_bytes_s'] = sum(int(r['output_bytes']) for r in rows) * 1e9 / sum(int(r['elapsed_ns']) for r in rows)
    for field in ['budget_after_work', 'budget_after_input', 'source_read_calls', 'source_requested_bytes', 'source_returned_bytes']:
        values = {int(r[field]) for r in rows}
        assert len(values) == 1, field
        result[field] = values.pop()
    return result

def delta(before, after):
    metrics = {}; flags = []
    for key, value in before.items():
        pct = (after[key] / value - 1) * 100 if value else None
        metrics[key] = {'before': value, 'after': after[key], 'delta_pct': pct}
        if pct is not None:
            if (key.startswith(('elapsed_', 'edit_')) or key == 'rss_kib') and pct > 5:
                flags.append(key)
            if key == 'throughput_output_bytes_s' and pct < -5:
                flags.append(key)
    return {'metrics': metrics, 'adverse_flags': flags}

comparisons = []
for paragraphs in [128, 512]:
    for count in [1, 8, 32]:
        for source in ['owned', 'file']:
            stem = f'p{paragraphs}-k{count}-{source}'
            states = {}
            for label, phase, mode in [('before_scalar', 'before', 'repeated'), ('after_scalar', 'after', 'repeated'), ('after_batch', 'after', 'batch')]:
                states[label] = load(phase, f'{stem}-{mode}')
            assert states['before_scalar'][2] == states['after_scalar'][2] == states['after_batch'][2], stem
            for kind, left, right in [('scalar_before_after', 'before_scalar', 'after_scalar'), ('after_batch_vs_scalar', 'after_scalar', 'after_batch'), ('batch_after_vs_scalar_before', 'before_scalar', 'after_batch')]:
                b, brss, _ = states[left]; a, arss, _ = states[right]
                before = stats(b); after = stats(a)
                before['rss_kib'] = brss; after['rss_kib'] = arss
                comparisons.append({'name': stem, 'kind': kind, 'paragraphs': paragraphs, 'replacements': count, 'source': source,
                                    'aggregate': delta(before, after), 'repeats': {
                                        str(i): delta(stats([r for r in b if int(r['repeat']) == i]), stats([r for r in a if int(r['repeat']) == i])) for i in [0, 1]}})
counts = {kind: dict(Counter(flag for c in comparisons if c['kind'] == kind for flag in c['aggregate']['adverse_flags'])) for kind in sorted({c['kind'] for c in comparisons})}
result = {'children': 36, 'measured_samples': 2160, 'warmup_samples': 216, 'comparisons': comparisons,
          'aggregate_flags_by_comparison_kind': counts, 'scope': 'Historical same-API scalar controls and separately labelled batch alternatives; whole-child RSS counted once; open/commit/publish/drop phases descriptive, lifecycle/edit/throughput/RSS adverse flags retained.'}
(HERE / 'comparison.json').write_text(json.dumps(result, indent=2) + '\n')
lines = ['# 0500 managed DOCX comparisons', '', 'Thirty-six children contain 2,160 measurements and 216 warmups. Every output identity agrees across matched routes. Work charges may differ as passes are removed. Percentiles use nearest rank; RSS includes whole-child setup and verification.', '']
for kind in counts:
    lines += ['## ' + kind, '', '| Workload | Lifecycle p50 before → after µs | Edit p50 before → after µs | Lifecycle p99 change | RSS change | Flags |', '| --- | ---: | ---: | ---: | ---: | --- |']
    for c in comparisons:
        if c['kind'] != kind: continue
        m = c['aggregate']['metrics']
        lines.append(f"| {c['name']} | {m['elapsed_p50_us']['before']:.2f} → {m['elapsed_p50_us']['after']:.2f} | {m['edit_p50_us']['before']:.2f} → {m['edit_p50_us']['after']:.2f} | {m['elapsed_p99_us']['delta_pct']:+.2f}% | {m['rss_kib']['delta_pct']:+.2f}% | {', '.join(c['aggregate']['adverse_flags']) or '—'} |")
    lines += ['', 'Aggregate flags: ' + json.dumps(counts[kind], sort_keys=True), '']
lines += ['Per-repeat values, all named phases, source counters, charged Work, and every retained flag appear in comparison.json. Two repeats on a shared host do not establish isolated causal costs or native/cold behavior.']
(HERE / 'comparison.md').write_text('\n'.join(lines) + '\n')
print(json.dumps({'comparisons': len(comparisons), 'aggregate_flags': counts}))
