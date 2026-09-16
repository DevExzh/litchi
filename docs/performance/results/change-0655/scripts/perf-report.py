#!/usr/bin/env python3
"""Summarize perf stat cycles and instructions across the four legs."""
import pathlib, sys
S = pathlib.Path('/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655')

def read(case, leg):
    out = {}
    for line in (S / 'out' / f'perf-{case}-{leg}.txt').read_text().splitlines():
        parts = line.split(',')
        if len(parts) > 2 and parts[0].replace('.', '').isdigit():
            out[parts[2]] = float(parts[0])
    return out

def main(case):
    legs = {leg: read(case, leg) for leg in ('A1', 'B1', 'B2', 'A2')}
    for leg in ('A1', 'B1', 'B2', 'A2'):
        v = legs[leg]
        print(f'leg\t{leg}\tcycles={v["cycles"]:,.0f}\tinstructions={v["instructions"]:,.0f}'
              f'\ttask-clock={v["task-clock"]:,.0f}ms')
    for metric in ('cycles', 'instructions', 'task-clock'):
        b = (legs['A1'][metric] + legs['A2'][metric]) / 2
        a = (legs['B1'][metric] + legs['B2'][metric]) / 2
        aa = 100 * (legs['A2'][metric] - legs['A1'][metric]) / legs['A1'][metric]
        print(f'{metric}\tbefore={b:,.0f}\tafter={a:,.0f}\tdelta={100*(a-b)/b:+.2f}%\tA/A={aa:+.2f}%')

if __name__ == '__main__':
    main(sys.argv[1])
