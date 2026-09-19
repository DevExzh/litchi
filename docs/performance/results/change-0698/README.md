# Change 0698 — borrow inherited MCE namespace views

performance_claim: none

Disposition: **rejected**. The production codec is restored to baseline;
three regression tests and the experiment remain. A repeated reachable
refusal regression exceeds the review threshold despite successful-path gains.
`candidate-codec.rs.txt`, `source-diff.patch` and `rejection.json` preserve and
bind the measured implementation. The timings do not describe retained gains.

The production candidate tests the narrow ownership hypothesis recorded in
[the final performance record](../../0698-mce-borrowed-inherited-namespaces.md):
the ephemeral Inherited view borrows namespace owners from the parent frame
while the child Ctx and emitted boundary remain owned. The source diff is
restricted to the shared MCE codec and its focused tests. See [design.md](design.md) and [code-review.md](code-review.md).

The packet uses the same 13-workflow PPTX probe and ten-case refusal probe as
0696, with new frozen binaries and current production sources. The shared MCE
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
python3 docs/performance/results/change-0698/prepare-control.py
python3 docs/performance/results/change-0698/check-control.py
python3 docs/performance/results/change-0698/build.py baseline
python3 docs/performance/results/change-0698/build-refusal.py baseline
python3 docs/performance/results/change-0698/build-oracle.py baseline
python3 docs/performance/results/change-0698/assembly.py baseline
python3 docs/performance/results/change-0698/ownership-assembly.py baseline
python3 docs/performance/results/change-0698/measure.py baseline
python3 docs/performance/results/change-0698/measure-allocations.py baseline
python3 docs/performance/results/change-0698/measure-refusal.py baseline
python3 docs/performance/results/change-0698/profile.py baseline
```

After the baseline build and AA measurements, apply only the `tests.rs` part of
the retained source diff and prove the baseline focused tests. Apply the
`codec.rs` part only after that proof, format the candidate, and record the
candidate focused receipt and final source diff:

```sh
git apply --include=crates/litchi-ooxml-common/src/mce/tests.rs \
  docs/performance/results/change-0698/source-diff.patch
python3 docs/performance/results/change-0698/run-focused.py baseline
git apply --include=crates/litchi-ooxml-common/src/mce/codec.rs \
  docs/performance/results/change-0698/source-diff.patch
cargo fmt --all
python3 docs/performance/results/change-0698/run-focused.py candidate
python3 docs/performance/results/change-0698/source-diff.py
```

Invoke the three build drivers, assembly.py and ownership-assembly.py with candidate. Run measure.py
and measure-refusal.py with `compare`, then measure-allocations.py and
profile.py with `candidate`. Run binary-sizes.py after both phases. Native AA
plus ABBA means four baseline and two candidate legs, each with 100 samples/five
warmups. Allocator diagnostics use separate instrumented binaries with three
samples.

After native profiling is complete, run the correctness oracle:

```sh
python3 docs/performance/results/change-0698/oracle/corpus.py \
  /path/to/litchi-0698-bin/baseline-oracle \
  /path/to/litchi-0698-bin/candidate-oracle \
  /path/to/checkout/test-data \
  /path/to/packet/oracle-results
python3 docs/performance/results/change-0698/measure-oracle-controls.py
python3 docs/performance/results/change-0698/measure-oracle-real-controls.py
```

Do not use the corpus driver's optional timing mode as native evidence. Its
invocation elapsed times are bookkeeping. The secondary native controls are
separate 300-sample AA/ABBA measurements with default, one-name and 4,096-name
extension profiles. Run them without overlapping Cargo or other profiling.

Generate summaries with summarize.py, summarize-allocations.py,
summarize-refusal.py and report-metrics.py. Run
measure-oracle-declaration-controls.py before the repository gates to price
the outlined nonempty helper on 1,000 redeclaring siblings and mixed inherited
children. All initial/native/oracle/declaration timing controls finish before
the integration and evidence gates:

```sh
python3 docs/performance/results/change-0698/run-integration.py
python3 docs/performance/results/change-0698/quality-summary.py
python3 docs/performance/results/change-0698/run-evidence.py
```

After those gates are terminal, run the isolated follow-up without overwriting
any initial matrix:

```sh
python3 docs/performance/results/change-0698/measure-followup.py
```

It records separate native, refusal, and opaque oracle raw files under
`followup/`, with ABBA 300-sample/10-warmup receipts, exact oracle identity
parity, recomputed summaries, comparisons, and triggers. After reviewing the
results, reproduce the recorded rejection with reject-candidate.py. It preserves
the candidate codec, restores only that production file to baseline and binds
both source states in rejection.json. Run run-focused.py with `retained` to
verify the final codec/test combination. Then run audit.py, cleanup.py --apply and
seal.py. The audit distinguishes the frozen measured candidate from the retained
baseline codec and regression tests. Keep the final script-hashes.json manifest
synchronized with every driver adaptation. Cleanup removes exactly the batch
target, frozen binary directory, raw profile directory, and generated
marker-control archive.

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

The main probe identifies as 0698 and intentionally retains edit marker 0691.
The refusal transcript keeps its original 0693-refusal identity; its binary is
probe0698. Their inherited sub-probe READMEs retain original provenance; this
README's commands govern this batch. The oracle identifies as 0698-mce-oracle.
