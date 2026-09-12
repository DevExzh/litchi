# 0521: borrow XLSX validation events and namespaces

This bundle compares the same source-backed scalar cell edit/save APIs before
and after removing per-event owned conversion and namespace-resolver cloning.
The [change record](../../changes/0521-xlsx-borrow-validation-events.md) gives
the results and admission decision. OLE2/OOXML remains the active priority;
ODF is deferred and iWork excluded.

## Reproduction and custody

`plan.json` freezes the hypothesis, exact matrix, CPU affinity and admission
criteria at revision `672b8468ac5693136cd27df62c541923485eb444`. Each stage has
its own source manifest, tracked source patch, build logs, binary descriptors
and terminal capture receipts. The differential test file is identical in
both manifests and retained in the final repository. `verify.py` reconstructs
both tracked snapshots through private Git indexes and checks every source
hash; only the production validation file differs between stages.

To reconstruct baseline code, start from the recorded revision, apply
`baseline/source.patch`, and copy `validation_borrow_tests.rs` from the final
change into the same source directory. Its hash must match the baseline
manifest. Applying `borrow-candidate.patch` then reconstructs the candidate.
Keep an external copy of this bundle when checking out the older revision.

Builds use the standalone benchmark Cargo workspace, its committed lockfile
and ordinary release defaults. The root workspace's release LTO settings do
not apply. The normal executable and separate `allocator-metrics` executable
have independent hashes. `host.json` retains the actual tool and machine
observations; `run.py` checks source and binary hashes around each operation.

For a new campaign use a fresh evidence directory and plan. The retained
driver intentionally refuses to overwrite existing stages or receipts:

```sh
python3 -B docs/performance/results/change-0521/run.py baseline freeze
python3 -B docs/performance/results/change-0521/run.py baseline build-normal
python3 -B docs/performance/results/change-0521/run.py baseline native
python3 -B docs/performance/results/change-0521/run.py baseline profile
python3 -B docs/performance/results/change-0521/run.py baseline build-alloc
python3 -B docs/performance/results/change-0521/run.py baseline alloc
```

Apply `borrow-candidate.patch` only after baseline capture, then use the same
six actions for `candidate`. Baseline already includes the harness allocation
observer and tests, as shown by its source patch. The production candidate
does not alter parser policy, validation ordering, budgets or candidate readback.

Recompute retained evidence without build binaries:

```sh
python3 -B docs/performance/results/change-0521/analyze.py /tmp/0521-comparison.json
python3 -B docs/performance/results/change-0521/analyze_profiles.py /tmp/0521-profiles.json
python3 -B docs/performance/results/change-0521/verify.py --sealed
```

Profile analysis regenerates annotations deterministically with
`PERL_HASH_SEED=0` and `PERL_PERTURB_KEYS=0`; it uses the retained 0519 raw
parser helpers whose hashes appear in its report. Only the fourth numbered
dump measures the selected one-percent commit. The first three lifecycle
commits and final zero-Ir dump remain in the bundle.

## Measurement boundaries

Each native stage retains 700 samples in 14 fresh children: four primary
children with 20 warmups/100 samples, and ten guards with 10 warmups/30 samples.
Total time is open + planning + staged sets/commit + publication. Publication
includes dropping its returned snapshot. Sink setup, other handle destruction,
reopen and correctness oracles are outside that sum. Phase vectors use
acquisition order; sorted totals are aligned through `sample_order`.

Each allocator stage retains 20 samples in four separate children with no
warmup. The canonical System observer covers staged sets and commit only;
allocator-build timings are excluded from native latency comparisons.
Incremental peak means region peak minus live bytes at entry. Whole-child RSS
includes setup, reopen and oracles and is a distinct measurement.

Input is the instrumented in-memory source with a fresh editor/cache each
iteration. Default provider options do not activate filesystem, cold-cache
or simulated range lanes. The synthetic vendor-extension fixture exercises
unknown package-member preservation, not unknown worksheet grammar. The
one-percent primary shapes touch all four sheets, so zero unselected-sheet
reads is not evidence of selective-subset access.

## Correctness and limitations

The 11 differential tests compare admission/refusal and exact error strings
with a frozen copy of the old loop, sharing unchanged policy helpers. The
first pre-capture baseline test run exposed one incorrect expected error in
the test fixture; its failed log and corrected 11-test pass are both retained.
Neither capture stage includes that incorrect fixture expectation.

`checks.py` records the serial correctness and policy commands. The analyzers
check phase sums, source/budget/corpus/output identities, allocation balances,
receipt custody and selected profile accounting. Four negative raw-vector
tests demonstrate rejection without relying on checksum failure.

No cold/range, native Office-producer, fuzz, hardware-counter, parallel-scaling
or broad coverage-catalog completion follows. Callgrind Ir is a diagnostic
guest-instruction count, not elapsed time or allocation count. Collection-off
call metadata cannot be interpreted as timed event counts. Remaining owned
names and stacks are intentional; this is not allocation-free validation.
