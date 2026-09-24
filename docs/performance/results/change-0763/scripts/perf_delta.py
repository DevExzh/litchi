#!/usr/bin/env python3
"""Per-iteration user-space instructions and cycles of one harness case.

Runs the binary twice with the same warmup and different sample counts under
`perf stat -x,`, and divides the counter difference by the sample difference.
Everything outside the timed loop (corpus construction, reopen and digest
gates, process start) is identical in both runs and cancels out. Each process
is pinned with taskset.
"""
import json, subprocess, sys, os, tempfile

def run(binary, case, shape, samples, warmup, core, tmpdir):
    out = tempfile.NamedTemporaryFile(dir=tmpdir, suffix='.csv', delete=False).name
    report = tempfile.NamedTemporaryFile(dir=tmpdir, suffix='.json', delete=False).name
    os.unlink(report)
    events = os.environ.get('PERF_DELTA_EVENTS', 'instructions:u,cycles:u')
    cmd = ['perf', 'stat', '-x,', '-o', out, '-e', events, '--',
           'taskset', '-c', str(core), binary, '--case', case, '--samples', str(samples),
           '--warmup', str(warmup), '--json', report]
    if shape:
        cmd[cmd.index('--case'):cmd.index('--case')] = []
        cmd += ['--semantic-shape', shape]
    env = dict(os.environ, TMPDIR=tmpdir)
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    counters = {}
    for line in open(out):
        parts = line.strip().split(',')
        if len(parts) > 2 and parts[0].replace('.', '').isdigit():
            counters[parts[2].split(':')[0]] = int(float(parts[0]))
    os.unlink(out)
    data = json.load(open(report))
    os.unlink(report)
    digest = sorted({r.get('output_sha256') or 'none' for r in data['results']})
    return counters, digest

def main():
    binary, case, shape, low, high, warmup, core, tmpdir = sys.argv[1:9]
    low, high, warmup, core = int(low), int(high), int(warmup), int(core)
    c_low, d_low = run(binary, case, shape, low, warmup, core, tmpdir)
    c_high, d_high = run(binary, case, shape, high, warmup, core, tmpdir)
    per = {k: (c_high[k] - c_low[k]) / (high - low) for k in c_low}
    print(json.dumps({'binary': binary, 'case': case, 'shape': shape, 'low': low, 'high': high,
                      'warmup': warmup, 'core': core, 'counters_low': c_low, 'counters_high': c_high,
                      'per_iteration': per, 'output_sha256': sorted(set(d_low) | set(d_high))}))

if __name__ == '__main__':
    main()
