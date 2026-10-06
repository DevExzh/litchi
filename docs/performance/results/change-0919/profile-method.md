# 0919 Heaptrack allocation attribution method

This is analysis of retained captures, not a modification to production or
benchmark source. The unchanged 0917 native binary is the only profiled
executable. All six probe JSON reports passed the existing 0917 validator with
`expected_warmups=0` and `expected_samples=1`.

The captured interpreted format starts with `v 10500 3`. Its string, instruction
and parent-trace tables are one-based; allocation-info entries are zero-based.
An `a` entry gives hexadecimal requested size and trace index. Each `+` event
names an allocation-info index. Counting those events and summing their sizes
attributes requested allocations, including replacement allocations on growth.
The `-` events are deallocation events and are not part of this requested-byte
sum. This method does not infer lifecycle peak-live memory from event totals.

For each trace, the parent relation propagates the presence of the exact
`doc_retention_probe_0917::measure_one` function and subtree function names.
Only allocation events whose trace contains that boundary are included.
Function memberships are inclusive and overlap. The exclusive partition in
`phase-results.json` applies its stated priority to each trace's joint
memberships; it sums to the same lifecycle totals. In particular, final commit
takes precedence over the nested `Snapshot::open_bounded` it invokes.

The following read-only decoder reproduces the scoped totals and inclusive
memberships for one retained recording. It checks the interpreted version,
table references and string lengths, and fails if a new record kind appears.
Repeat it for every path in `profile-results.json`. The results must match that
record and the 0917 allocator totals, not merely its own second execution.

```sh
python3 -B - /tmp/litchi-performance-0919/r01-docfloat-e32/heaptrack.zst <<'PY'
import collections
import subprocess
import sys

tokens = (
    'parse_sprms', 'rewrite_chpx', 'source_paragraphs',
    'capture_container', 'read_stream_from_fat', 'read_stream_from_minifat',
    'main_story_text', 'commit_candidate_with_rendered',
    'render_with_layout', 'plan_reuse',
)
strings, instructions, traces, allocations = [''], [()], [frozenset()], []
totals, memberships = collections.Counter(), collections.Counter()
version_seen = False
process = subprocess.Popen(['zstdcat', sys.argv[1]], stdout=subprocess.PIPE,
                           text=True, encoding='utf-8')
for line in process.stdout:
    tag, _, rest = line.rstrip('\n').partition(' ')
    if tag == 'v':
        assert not version_seen and rest == '10500 3'
        version_seen = True
    elif tag == 's':
        length, _, value = rest.partition(' ')
        assert len(value.encode()) == int(length, 16)
        strings.append(value)
    elif tag == 'i':
        fields = rest.split()
        instructions.append(tuple(strings[int(fields[i], 16)]
                                  for i in range(2, len(fields), 3)))
    elif tag == 't':
        instruction, parent = (int(x, 16) for x in rest.split())
        assert instruction < len(instructions) and parent < len(traces)
        traces.append(traces[parent] | frozenset(instructions[instruction]))
    elif tag == 'a':
        size, trace = (int(x, 16) for x in rest.split())
        assert trace < len(traces)
        allocations.append((size, trace))
    elif tag == '+':
        size, trace = allocations[int(rest, 16)]
        names = traces[trace]
        if not any('doc_retention_probe_0917::measure_one' in n for n in names):
            continue
        totals['calls'] += 1
        totals['bytes'] += size
        for token in tokens:
            if any(token in name for name in names):
                memberships[token, 'calls'] += 1
                memberships[token, 'bytes'] += size
    else:
        assert tag in {'X', 'I', '-', 'c', 'R', '#', ''}, tag
assert process.wait() == 0 and version_seen
print(dict(totals))
for token in tokens:
    print(token, memberships[token, 'calls'], memberships[token, 'bytes'])
PY
```

For the displayed FloatingPictures/32 recording the required totals are
95,057 calls and 164,233,879 bytes. For FloatingPictures/two they are 18,014
and 17,369,691; for NoHeadFoot/32 they are 19,024 and 16,781,690. Each pair of
repetitions must agree independently.

The second representation is produced by the exact retained command:

```text
heaptrack_print -f RECORDING --filter-bt-function measure_one -p 0 -T 0 -a 1 -n 20 -s 10 --flamegraph-cost-type allocations -F FRESH_STACK_FILE
```

Every nonempty folded-stack line must contain the measured boundary. Summing
the decimal integer after its final space must reproduce the call count above.
The two folded-stack files per workload have identical SHA-256 values despite
different raw recording addresses/timestamps. This cross-check does not turn
the inclusive call counts into native-time percentages.

For the top-byte trace list, aggregate requested size and count by raw trace
index before sorting by descending bytes. Walk that trace's instruction/parent
chain and retain the `litchi_` library frames for inspection. These trace IDs
are local to `r01-docfloat-e32`; they are not stable addresses or identifiers
for another recording. `top-byte-traces.json` intentionally retains full
library chains, so both output-emission stacks can be distinguished from
package recapture and mutation even when their leaf owner has the same name.

## Limits and review status

The decoder above has been reconciled with a different System allocator
instrumentation lane and Heaptrack's own folded-stack printer on all six
captures. It has not yet received the requested independent agent review.
No corruption-testing or independently reviewed portable replay is claimed.
The raw captures, analysis commands, protocol and reduced data are retained
for that next audit. No production implementation is authorized solely by
these diagnostic results.
