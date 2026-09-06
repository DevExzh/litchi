# 0438: ODP markup batching negative result

The candidate was rejected and production restored. Medium/large p50 improved
0.972–1.982%, below the immutable 5% gate in `before-hypothesis.json`.
See [the change record](../../changes/0438-odp-markup-batching-negative.md).

`protocol.json` is the frozen measurement protocol. `before/build.json` and
`after/build.json` bind exact executable hashes, source manifests, release
commands and toolchain. Both builds have the same Git HEAD; their manifests
differ only in `streaming.rs` and `streaming/markup_tests.rs`. The candidate
sources are stored as `.txt` so they do not enter later Rust source manifests.
`candidate/candidate.patch` applies to the recorded baseline and includes the
seven tests. No candidate production code is retained in the crate.

The matrix is 24 reports / 720 samples, CPU 2, one worker, 64/4,096/8,192 slides,
normal/allocator modes, 30 samples / three warmups, A1/B1/B2/A2 order. Twelve
preparatory pilot reports and four fresh formal profiles are separate.
Historical 0437 artifacts under `prior-profile` support only the initial
hypothesis: their source manifest and binary exactly match the fresh baseline.
They are not fresh 0438 captures. All report semantic checks use the unchanged
0437 oracle mapped to its `after-streaming` role.

`summary.py --check` rederives the formal summary. `decision.py --check`
rederives the retention result. Raw allocator vectors are retained; equality
compares chronologically aligned operation counters, live deltas and region
peak above entry. Absolute process live/high-water values include path storage.
`draft-history` retains the initial adapter errors and their corrections;
the failed preflight remains in `checks`. No measurement was discarded.

The 66 matched comparisons and repeat checks produce no flags above 5%.
There is consequently no flagged regression requiring acceptance; the
candidate is rejected on insufficient practical gain. Source review and
ADR boundaries are recorded in `source-review.md` and `adr-compliance.md`.

After sealing, replay from an exported bundle with:

```sh
python3 -B verify.py --portable-check --require-inventory --stage final
python3 -B lifecycle.py --stage final
python3 -B summary.py --check
python3 -B decision.py --check
```

Fresh reruns require rebuilding and copying the two recorded implementations
and using a new evidence directory. Retained successful receipts and formal
capture paths are immutable. Copied executables and this batch's external
draft directories are removed only after portable replay; both shared Cargo
targets and the untracked user goal are preserved.

The restored baseline passed 349 ODP tests. `source-restoration` proves its
source manifest equals the original baseline and the archived candidate patch
applies. Eight copied-bundle mutation probes pass. Cleanup removed exactly
five task directories totaling 1,827,055,009 bytes. `lifecycle-contract.json`
binds these later cleanup/replay details without changing the frozen capture
protocol. Historical failed adapter/replay checks remain visible.
