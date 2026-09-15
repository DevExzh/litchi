# change-0580 evidence packet: a target-scoped ZIP strict-layout proof

Change record:
[`docs/performance/0580-zip-target-scoped-strict-layout.md`](../../0580-zip-target-scoped-strict-layout.md).
Design record: [`0575`](../../0575-zip-lazy-strict-layout-design.md), whose
candidate (b) this implements.
Disposition: retained. `performance_claim: none`; **no timing is measured and
none is claimed.** This change carries a **semantic difference**, stated in the
record's "The semantic difference, in bytes" section.

## Contents

| Path | What it is |
| --- | --- |
| `make_strict_scope_probe.py` | The throwaway instrumentation that produced the counts. It inserts a counting `ReaderAt` and a scenario walk into `crates/soapberry-zip/src/office.rs`'s test module, because `IndexedArchive::strict_layout_for` is private and is exactly the path being measured. It calls only APIs that exist on both sides of the change, so one file measures both trees. |
| `scope-before.txt`, `scope-after.txt` | Reads and bytes the strict-layout path alone issues, per fixture and per scenario, before and after. Scenarios: open-and-list, one member, a one-part closure, every member in physical order, and every member in reverse order. |
| `census-before.txt`, `census-after.txt` | Per-fixture accept/refuse verdict string, read count and byte count for reading **every** member of every DOCX/PPTX/XLSX fixture under `test-data/ooxml` — 168 fixtures, 4,089 members. `diff` between the two files is **empty**. |
| `prechange-tests.txt` | The eleven new tests run against the pre-change tree, and the reason each of the six failures fails. |
| `residual_window.py` | Re-derives change 0575's residual window, prices both brackets per target, and models the per-record memo in read order. Pure Python over the fixture bytes; no build required. |
| `residual-window.txt`, `residual-window.json` | Its output. |

## Result

Reads issued by the strict-layout path alone, post-0573 tree → this change:

```
sheet-names.xlsx (13)                   one member  13 ->  8     closure  13 -> 11
ConditionalFormattingSamples.xlsx (132) one member 132 -> 20     closure 132 -> 59
shapes.pptx (48)                        one member  48 -> 15     closure  48 -> 16
universal-content.xlsx (13, all desc)   one member  39 ->  9     closure  39 -> 36
comment.docx (10, all desc)             one member  30 -> 11     closure  30 -> 11
shape-soft-edges.pptx (65, all desc)    one member 195 -> 65     closure 195 -> 86

open and list:                            0 ->   0   on every fixture
every member, physical order:  identical to the archive-wide proof on all 168
                               fixtures: 4385 reads, 517371 bytes, 4089 accepted
every member, reverse order:   13 -> 25, 132 -> 263, 48 -> 95, 65 -> 323
```

Change 0575's frozen prediction scores **9 of 12 read predictions confirmed**.
The three falsified ones have two independent causes, both reproduced by
`residual_window.py`:

1. 0575's residual window of 65,559 bytes assumes a predecessor's local name
   length equals its central name length. Candidate (b) never validates a
   predecessor's name, so both `u16` halves of the local variable region are
   unknown and the sound window is 131,094 bytes. On the only fixture large
   enough for the window to matter, one member goes from a predicted 4 reads to
   a measured 20, and the closure from 14 to 59.
2. 0575's cost model counts the union of records each target touches; the
   implementation memoises per record, so a record probed as a neighbour and
   later read as a target costs two reads. `shapes.pptx`'s closure is 16, not
   15.

## Replay

Recompute the scenario comparison from the retained files:

```sh
python3 -B -c "
import re, pathlib
def load(p):
    return {(m.group(1), m.group(2)): (int(m.group(3)), int(m.group(4)))
            for m in re.finditer(r'SCOPE (\S+) scenario=(\S+) reads=(\d+) bytes=(\d+)',
                                 pathlib.Path(p).read_text())}
b = load('docs/performance/results/change-0580/scope-before.txt')
a = load('docs/performance/results/change-0580/scope-after.txt')
for k in b:
    print(f'{k[0]:<45} {k[1]:<12} {b[k][0]:>4}r/{b[k][1]:>6}B -> {a[k][0]:>4}r/{a[k][1]:>6}B')"
```

Confirm convergence:

```sh
diff docs/performance/results/change-0580/census-before.txt \
     docs/performance/results/change-0580/census-after.txt && echo identical
```

Re-derive the residual window and the ordered read model:

```sh
python3 docs/performance/results/change-0580/residual_window.py
```

To reproduce the capture, extract the committed revision twice, overlay the
change on one of them, insert the probe into both, and run each with its own
`CARGO_TARGET_DIR`:

```sh
git archive 32d25e088 | tar -x -C "$SCRATCH/before"
cp -a "$SCRATCH/before" "$SCRATCH/after"
cp crates/soapberry-zip/src/{archive.rs,office.rs} "$SCRATCH/after/crates/soapberry-zip/src/"
for tree in before after; do
  python3 docs/performance/results/change-0580/make_strict_scope_probe.py \
      "$SCRATCH/$tree/crates/soapberry-zip/src/office.rs"
  (cd "$SCRATCH/$tree" && CARGO_TARGET_DIR="$SCRATCH/target-$tree" \
      cargo test -p soapberry-zip --lib zz_target_scoped -- --nocapture --test-threads=1)
done
```

The extraction is not optional. Change 0572 had to discard a whole capture
because a probe with path dependencies on the shared working tree linked another
change that was in flight, with no warning.

Note that `cargo test --nocapture` prints the first line of test output on the
same line as the test name, so an extraction anchored at the start of a line
silently drops the first record. The retained files were extracted with
`grep -oE "SCOPE .*"` and `grep -oE "CENSUS[A-Z]* .*"`.

## What is not here

No timing, allocation-profile, cold-cache, physical-device or cross-platform
capture. The counts are logical `ReaderAt` calls over warm fixtures. No
range-source arm was run for this change: change
[0572](../../0572-ooxml-range-source-attribution.md) is the record that prices
these requests against a latency-bearing source, and it is the measurement that
satisfied change 0575's admission gate.

The `parse_zip` fuzz target was not run — `cargo-fuzz` is not installed in this
environment and no nightly toolchain is present. That gate is outstanding.
