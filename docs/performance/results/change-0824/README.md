# 0824 — outer-whitespace compaction proof trial

**Adopted.** The admitted real-file public edit improves paired median p50 by
10.407% (ratio 0.895935; CI95 0.892216–0.899349), saving 298 allocation calls
and 21,688 allocated bytes. All nineteen rows pass latency/allocation guards.
Large synthetic capture is 3.636% slower; no full-save, synthetic-wide, tail,
or RSS saving is claimed. See [the report](../../0824-pptx-outer-whitespace-compaction.md).

Final validation and independent replay pass for 342 reports/7,106 samples.
The original candidate-quality failure was a new cold-path test expecting
zero reads instead of the required one. `pre-repair-0` and `quality-after-0`
preserve the complete original evidence; `failure_audit.py` verifies that only
that test assertion/message changed before a fresh baseline freeze/build and
qualification. No comparative capture preceded recovery. All corrected
quality gates pass (1,253 candidate tests; three ignored), and no offline
reader invocation failed. Cleanup verifies eight admitted binaries and removes
9,686 files/3,238,494,566 logical bytes; the earlier baseline-only cleanup is
separately recorded. Unrelated workspace files are unchanged.

This packet tests a transient exact-root-byte proof in the PPTX opened
transaction/XML owners. The hypothesis is that removing whitespace outside an
otherwise byte-identical document root need not trigger two additional full
Scene parses when the staged payload already passed that read. Any changed
root or unavailable proof takes the existing comparison path. Production APIs,
written bytes, initial validation, final capture, and durability are unchanged.

The base is `7b268927bfe3ed9c9000bfd567fc4ce656ec0cd9`; the only candidate
production paths are `opened/transaction.rs` and `opened/xml.rs` in
`crates/litchi-pptx/src`. Source/profile basis, candidate archive, independent
review, and exact frozen inputs accompany the execution receipts. The raw 0822
profile remains historical call-path evidence, not a new timing measurement.

## Protocol

Both complete probe trees are copied from the final 0823 sources, including its
pre-freeze allocator-test fixes. Only the packet path, report schema, and base
identity change. Dependency locks and timed corpus/operation bodies stay the
same. Synthetic outputs remain pinned to the sealed historical qualification
oracles; the real file must exactly reproduce the independently admitted 0821
output, complete reopened text, target text, and slide count.

Fresh baseline and candidate production quality run six serial commands each:
fmt, all-feature/all-target check, full all-feature tests, warning-denied Clippy,
warning-denied rustdoc, and crate boundaries. Probe quality runs eighteen
commands across native and allocation features. Each failed invocation must
retain its original log, receipt, and source before any correction.

Freeze binds the complete production and ordinary-save harness source sets,
all 35 normative inputs, root/tool/probe locks, both probe trees, the candidate
before/after archive and combined patch, host/toolchain/corpus metadata, the
six driver files, and the three unrelated workspace-file hashes. Quality now
includes `.cargo/config.toml` directly in its source census.

Four release binaries per leg cover synthetic/real and native/allocation.
Builds use locked offline Cargo, two jobs, opt-level 3, debug level 1, thin LTO,
one codegen unit, unwind, and incremental compilation disabled. Root owns all
execution and Git mutations. Every child workload runs serially on CPU 12 with
an independent `/usr/bin/time` RSS receipt. There is no exclusive-host claim.

- Qualification: nineteen workflows per leg, one observer sample each,
  38 reports/38 samples.
- Native: six counterbalanced blocks, thirty measured samples and three
  warmups per process, 228 reports/6,840 samples.
- Allocation: two counterbalanced blocks, three measured samples and no
  warmup, 76 reports/228 samples. Observer elapsed times are not native timing.

Total: 342 reports/7,106 measured samples. The nineteen rows are six synthetic
fixtures crossed with capture, commit, and lifecycle, plus the admitted real
direct public edit. Synthetic capture is a negative control; synthetic
lifecycle includes in-memory serialization, while real edit excludes open,
serialization, readback, package-owner destruction, and any path-save/fsync.
The returned publication snapshot is dropped inside the timed real-edit helper.

The frozen benefit rule requires a paired median p50 ratio at most 0.97 with
upper bootstrap endpoint below one in at least one of the thirteen eligible
commit/lifecycle/real rows. Any of nineteen rows with lower endpoint above
1.05 vetoes adoption. Allocation calls/bytes/net-live/peak-above-entry cannot
increase. An RSS increase above 5% requires explicit review. Bootstrap uses
10,000 resamples, seed 824824, endpoints 250 and 9749. All rows, p50/p95/p99/
mean/RSS, and spread flags are exposed; no aggregate hides regressions.

## Execution and replay

The root-owned sequence on a fresh checkout of the pinned base is:

```text
python3 -B docs/performance/results/change-0824/prepare.py
python3 -B docs/performance/results/change-0824/quality.py before
python3 -B docs/performance/results/change-0824/fixture_basis.py --write
python3 -B docs/performance/results/change-0824/profile_basis.py --write
python3 -B docs/performance/results/change-0824/quality.py probes
# Review and format the archived candidate, then freeze.
python3 -B docs/performance/results/change-0824/freeze.py
python3 -B docs/performance/results/change-0824/build.py before
python3 -B docs/performance/results/change-0824/capture.py qualification before
# Apply the exact frozen combinedcandidate.patch.
python3 -B docs/performance/results/change-0824/quality.py after
python3 -B docs/performance/results/change-0824/build.py after
python3 -B docs/performance/results/change-0824/capture.py qualification after
python3 -B docs/performance/results/change-0824/capture.py native
python3 -B docs/performance/results/change-0824/capture.py allocation
python3 -B docs/performance/results/change-0824/analysis.py --write
python3 -B docs/performance/results/change-0824/raw_audit.py --write
python3 -B docs/performance/results/change-0824/validate.py
# Record disposition, review RSS if needed, and restore before source if rejected.
python3 -B docs/performance/results/change-0824/cleanup.py
python3 -B docs/performance/results/change-0824/validate.py --final
```

Drivers deliberately refuse existing output directories. Use fresh output
locations and the frozen source inputs to reproduce a trial; do not rerun the
write commands over a completed packet. Analysis and audit `--check`, both
basis `--check` commands, `validate.py --final`, and `seal.py --check-head` are
the offline replay routes once the record is complete. Cleanup witnesses retain
all eight exact binary identities after removing the owned target.
