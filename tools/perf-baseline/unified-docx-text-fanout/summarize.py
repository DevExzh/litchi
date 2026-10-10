"""Summarize repeated paired invocations without pooling away run variation."""
import json
import sys
from pathlib import Path


def summarize(paths):
    runs = [json.loads(Path(path).read_text()) for path in paths]
    if len(runs) < 2:
        raise ValueError("at least two independent invocations required")
    rows = []
    reference = runs[0]
    for run in runs:
        assert run['schema'] == 1 and run['order'] == 'ABBA'
        assert run['rounds'] >= 3
        assert run['leaf_routes'] == reference['leaf_routes']
        assert run['unified_routes'] == reference['unified_routes']
        assert [r['targets'] for r in run['results']] == [1, 10, 100, 1000, 10000]
    for index, baseline in enumerate(reference['results']):
        for mode in ('eager', 'source'):
            ratios, native, unified = [], [], []
            for run in runs:
                result = run['results'][index]
                assert result['corpus_sha256'] == baseline['corpus_sha256']
                assert result['corpus_bytes'] == baseline['corpus_bytes']
                value = result[mode]
                assert value['targets'] == baseline['targets']
                if baseline['targets'] == 10000:
                    assert value['status'] == 'refused_native_composite_replacement_limit'
                    assert value['timed'] is False
                    continue
                assert value['status'] == 'ok' and value['byte_equivalence']
                medians = []
                for key in ('leaf_ns', 'unified_ns'):
                    samples = value[key]
                    assert len(samples) == run['rounds'] * 2
                    assert all(isinstance(n, int) and n > 0 for n in samples)
                    medians.append(sorted(samples)[len(samples) // 2])
                assert medians == [value['leaf_median_ns'], value['unified_median_ns']]
                ratio = medians[1] / medians[0]
                assert abs(ratio - value['unified_over_leaf']) < 1e-12
                native.append(medians[0])
                unified.append(medians[1])
                ratios.append(ratio)
            rows.append(dict(targets=baseline['targets'], mode=mode,
                             status=baseline[mode]['status'],
                             leaf_medians_ns=native, unified_medians_ns=unified,
                             unified_over_leaf=ratios))
    return dict(schema=1, invocations=[Path(p).name for p in paths],
                statistic='upper median per invocation; no pooled acceptance claim',
                results=rows)


if __name__ == '__main__':
    print(json.dumps(summarize(sys.argv[1:]), indent=2))
