# Change 0696 — skip empty MCE namespace installation

performance_claim: none

The production candidate guards a namespace installation whose empty branch
only clones the current owner and immediately replaces that owner. See
[design](design.md), [code review](code-review.md), and the final
[performance record](../../0696-mce-empty-namespace-installation.md).

The packet uses the same 13-workflow PPTX probe and ten-case refusal probe as
0694, with new frozen binaries and current production sources. The shared MCE
oracle is built before and after the edit with one final lockfile, without
source restoration or dependency refresh. `baseline.json` binds the current
parent commit and all 33 previously read goal/ADR constraints.

## Reproduce

Use disposable baseline and candidate checkouts with this packet at the same
relative location. Preserve original results elsewhere: the measurement drivers
write new output. Start at the revision recorded in baseline.json; copy the
retained workspace-Cargo.lock to the disposable checkout root. Each standalone
probe has its own lockfile. Use the recorded Rust toolchain and dependencies.
CPU 12 and sibling scratch paths are machine-specific; adapt and record them
when reproducing on another host.

Before applying the production and test changes:

```sh
python3 docs/performance/results/change-0696/prepare-control.py
python3 docs/performance/results/change-0696/check-control.py
python3 docs/performance/results/change-0696/build.py baseline
python3 docs/performance/results/change-0696/build-refusal.py baseline
python3 docs/performance/results/change-0696/build-oracle.py baseline
python3 docs/performance/results/change-0696/assembly.py baseline
python3 docs/performance/results/change-0696/measure.py baseline
python3 docs/performance/results/change-0696/measure-allocations.py baseline
python3 docs/performance/results/change-0696/measure-refusal.py baseline
python3 docs/performance/results/change-0696/profile.py baseline
```

Apply the final codec/tests diff, run focused MCE tests, and invoke the three
build drivers and assembly.py with `candidate`. Run measure.py and
measure-refusal.py with `compare`, then measure-allocations.py and profile.py
with `candidate`. Run binary-sizes.py after both phases. Native AA plus ABBA
means four baseline and two candidate legs, each with 100 samples/five warmups.
Allocator diagnostics use separate instrumented binaries with three samples.

After native profiling is complete, run the correctness oracle:

```sh
python3 docs/performance/results/change-0696/oracle/corpus.py \
  /path/to/litchi-0696-bin/baseline-oracle \
  /path/to/litchi-0696-bin/candidate-oracle \
  /path/to/checkout/test-data \
  /path/to/packet/oracle-results
python3 docs/performance/results/change-0696/measure-oracle-controls.py
python3 docs/performance/results/change-0696/measure-oracle-real-controls.py
```

Do not use the corpus driver's optional timing mode as native evidence. Its
invocation elapsed times are bookkeeping. The secondary native controls are
separate 300-sample AA/ABBA measurements with default, one-name and 4,096-name
extension profiles. Run them without overlapping Cargo or other profiling.

Generate summaries with summarize.py, summarize-allocations.py,
summarize-refusal.py and report-metrics.py. Run run-integration.py,
quality-summary.py and run-evidence.py. After all Cargo and repository gates
finish, run measure-oracle-declaration-controls.py to price the outlined
nonempty helper on 1,000 redeclaring siblings and mixed inherited children.
Finally run audit.py, cleanup.py --apply
and seal.py. Keep the final script-hashes.json manifest synchronized with any
driver adaptations. The cleanup removes exactly the batch target, frozen binary
and raw profile directories plus the generated marker-control archive.

## Interpretation

Main native timers cover capture, working clone, text editing, commit and apply
on fresh prepared packages. File loading, initial opening, target search,
serialization and semantic/preservation oracles are outside. Perf's prefix
probe includes opening and capture and therefore has a different denominator.
The 10/210-iteration counter slope removes a constant startup estimate; it is
not a cold-cache or production RSS bound. Allocation requested-byte totals
include realloc replacement sizes and must not be added to realloc bytes again.

All timed work runs on a warm shared host with no quiescence claim. Raw samples,
per-leg bootstrap intervals and every >5% trigger remain explicit. The
same-length marker-stripped archive is a mechanism counterfactual, not a valid
document edit or general semantic-equivalence oracle. No new native Office
interoperability or CRUD-coverage claim follows from the timing probe.

The main probe identifies as 0696 but intentionally retains edit marker 0691.
The refusal transcript keeps its original 0693-refusal identity; its binary is
probe0696. Their inherited sub-probe READMEs retain original provenance; this
README's commands govern this batch. The oracle identifies as 0696-mce-oracle.
