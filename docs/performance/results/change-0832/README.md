# 0832 — defer dense column-map storage until the second record

The root-owned driver measures the four-file adaptive `Assignments` candidate
against the exact 0831 after revision. Candidate sources are archived at their
full repository paths under `candidate/before/` and `candidate/after/`.

The completed 0831 after quality run is reused as the 0832 before quality
witness only after its eight terminal receipts, logs, source hash, and
normalized full input census pass. The new packet still builds both binaries
and runs 18 one-sample qualifications for the before leg. The candidate
after leg receives a fresh eight-gate quality run, both builds, and both
qualification lanes. Comparative capture is refused until the independent
qualification admission records both reader hashes and all 36 report hashes.

The historical stages ran serially through the immutable launcher. Their
write-once receipts remain retained; replay commands for the completed packet
are listed below.

```sh
python3 -B docs/performance/results/change-0832/run_stage.py prepare
python3 -B docs/performance/results/change-0832/run_stage.py before
python3 -B docs/performance/results/change-0832/run_stage.py install
python3 -B docs/performance/results/change-0832/run_stage.py after
python3 -B docs/performance/results/change-0832/run_reader.py qualify.py
# Generate and preflight the supplemental fixtures/reader, then freeze its
# protocol and both binary legs before starting this comparative capture.
# See promotion-guard/README.md for the retained launcher commands.
python3 -B docs/performance/results/change-0832/run_guard.py driver.py prepare
python3 -B docs/performance/results/change-0832/run_stage.py capture
python3 -B docs/performance/results/change-0832/run_reader.py analyze.py
python3 -B docs/performance/results/change-0832/run_reader.py audit.py
```

The fixed matrix has nine cases, six counterbalanced native blocks, two
allocator-observer blocks, 180 retained reports, and 20,304 samples
including qualification. Native timing and observer allocation evidence are
reported separately. The requested-byte adoption gate is at least 50% for the
real edit, with identical output and review of native p50, RSS, allocation,
and entry-adjusted region-peak flags.

A separate [promotion guard](promotion-guard/README.md) measures two generated
multi-record worksheets. Its 48 reports and 12,040 samples are additional to
the unchanged parent matrix. Its qualification reader must admit the output
and allocation evidence before supplemental capture, and its regression flags
feed the root adoption decision. The source-extracted declaration diagnostic
in `layout/` measures Rust type sizes only; it supports no cache-miss claim.

Every child command has immutable start, terminal, and log receipts. Receipts
retain the source leg, binary leg, primary source hash, and four-file source
map. Failed commands remain in place; rerunning a stage requires a new packet.

The accepted candidate reduces real-edit requested bytes from 2,373,966 to
276,494 (88.35%) and paired native p50 by 10.40%. Both promotion fixtures have
unchanged allocation totals and entry-adjusted peaks, with no frozen regression
flag. See the [report](../../0832-xlsx-inline-column-map.md),
[decision](decision.json), and [results review](results-review.md) for every
case, intervals, diagnostics, and limitations. The two guard admission failures
are retained; its additive correction supplies only two missing globals and
changes no frozen reader code or statistic.

Cleanup removed the isolated target and both marked scratch directories.
Main analysis, independent audit, corrected guard analysis, and the adoption
decision all replay after cleanup. Replay without creating new packet files:

```sh
python3 -B docs/performance/results/change-0832/analyze.py --check
python3 -B docs/performance/results/change-0832/audit.py --check
python3 -B docs/performance/results/change-0832/promotion-guard/replay_reader_v2.py --check
python3 -B docs/performance/results/change-0832/decision.py
python3 -B docs/performance/results/change-0832/layout.py --check
python3 -B docs/performance/results/change-0832/census.py --check
python3 -B docs/performance/results/change-0832/seal.py verify
```
