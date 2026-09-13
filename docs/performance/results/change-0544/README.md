# 0544: XLSX event-bound preflight rejected

The candidate is rejected and production restored. The conservative preflight
avoids discarding over-cap parser state, but residual scan overhead still fails
two valid cap comparisons. One medium primary workflow repeat also misses the
required 3% p50/mean gain. No runtime speedup is retained.

The measured cap benchmark is retained as
`crates/litchi-xlsx/examples/perf_cap_boundary.rs`. It provides deterministic
stored XLSX fixtures, isolated planning timing, exact event counts, snapshot
checks and byte-exact no-op publication. Both baseline and candidate built the
same example; no Cargo dependency or production API changed.

- [Change record](../../changes/0544-xlsx-event-preflight-rejected.md)
- [Plan](plan.json), [protocol](protocol.md), [decision](decision.json)
- [Candidate design](design.md), [independent source review](source-review.md)
- [Native](comparison.json), [allocation](allocation-analysis.json), [refusal](guard-analysis.json)
- [Cap analysis](cap-boundary/cap-analysis.json), [individual reviews](adverse-review.json)
- [Quality](quality-summary.json), [independent results review](results-review.md), [next priority](next-priority.md)

Every main and cap capture completed. Main planning p50 improves 18.2–20.4%,
and workflow p50 improves 2.86–6.57%. The medium repeat2 p50/mean gains are
2.857/2.917%, below the fixed 3% requirement. Cap 164 repeat1 p50/mean regress
5.449/5.358%; cap 256 repeat2 regress 6.940/6.887%, above the fixed 5% limit.
Allocation and late-error envelopes pass, but same-invalid late-validator p50
increases 182.7–188.2% and allocated bytes increase 616–961%. Passing the
baseline-valid refusal envelope does not mean unchanged invalid-input cost.

The candidate passed 1,305 tests and Clippy before measurements. Its new oracle
covers direct NsReader event variants, references, malformed input, exact cap
and cap+1, plus conservative false-positive fallback. Candidate-only tests and
production changes remain preserved in snapshots/patches, not retained runtime.
Conditional profile/hardware/eager lanes were not executed after pilot failure.
Their prepared scripts make no performance claim.

The initial benchmark Clippy failure and interrupted baseline release build
are preserved. The interrupted attempt has partial logs and no terminal receipt;
a process/session check established it was stopped before the same frozen build
resumed. No exit status or completed binary is fabricated for that attempt.

`python3 -B verify.py --strict` replays the sealed evidence after cleanup.
Precleanup/preseal checkpoints preserve their then-pending status; the final
seal and strict replay prove completed custody. Captures are immutable and
historical drivers refuse overwriting receipts. The owned target and ten
executable copies are removed after hash-bound checks.

OLE2/OOXML remains active; ODF optimization is deferred until that goal completes.
