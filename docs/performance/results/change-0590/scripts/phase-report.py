#!/usr/bin/env python3
"""Per-phase medians for the PPTX boundary selectors, paired across the four legs."""
import json, pathlib, statistics, sys
S = pathlib.Path('/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0590')
PHASES = ('plan_ns', 'commit_ns', 'publication_ns', 'reopen_ns')

def med(value):
    if isinstance(value, dict):
        return value.get('p50')
    if isinstance(value, list):
        return statistics.median(value)
    return value

def main(case):
    rows = {}
    for leg in ('A1', 'B1', 'B2', 'A2'):
        data = json.loads((S / f'time-{case}-{leg}.json').read_text())
        row = next(r for r in data['results'] if r['case'] == case)
        summary = row['source']['pptx_slide_boundaries']
        rows[leg] = {phase: med(summary[phase]) for phase in PHASES}
        print('leg\t' + leg + '\t' + '\t'.join(f'{p}={rows[leg][p]/1e3:.1f}us' for p in PHASES))
    for phase in PHASES:
        b = (rows['A1'][phase] + rows['A2'][phase]) / 2
        a = (rows['B1'][phase] + rows['B2'][phase]) / 2
        print(f'phase\t{phase}\tbefore={b/1e3:.1f}us\tafter={a/1e3:.1f}us\tdelta={100*(a-b)/b:+.2f}%')

if __name__ == '__main__':
    main(sys.argv[1])
