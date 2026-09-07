# ODP publication audit proof reuse

The candidate is **retained** under the frozen gate. Normal owned append
lifecycle p50 improves 6.964% / 7.305% on medium and 7.366% / 6.820% on large
sources in R1/R2. All four required rows exceed the 3% threshold and have
negative independent bootstrap upper bounds. Tiny improves 10.049% / 8.968%.
No elapsed/RSS result exceeds the +5% review threshold and no allocator metric
increases. This is a scoped ODP result; the full non-iWork goal remains open.

## Mechanism and preservation

The private serializer returns an audit proof with one successfully finished
and reopened package. It binds strong `Arc<Vec<u8>>` owners for both the exact
source archive and candidate archive. The writer has already audited authored
core XML and XML-classified media under its strict default per-part limits.
Other XML is exact-source copied. An equivalent source manifest is preserved
verbatim; a generated manifest must pass the writer audit.

Checked accounting sums actual authored XML lengths and counts plus a
conservative 32 MiB/one-part manifest allowance. The proof checks all six
per-part limits and the existing 128 MiB/65,536-part aggregate bounds. Overflow,
missing source, source/candidate identity mismatch, ineligible accounting or a
later design/annotation/RDF/chart/content package replacement falls back to the
existing validator. This guard controls optimization eligibility, not admission.

Only the final `validate_compact_xml_parts` pass is skipped on a proof hit.
Writer validation, candidate ZIP reopening and full slide readback, the raw
source-reference precheck, media/domain readbacks, final snapshot selection,
exact no-op behavior and source-checked reversible patches retain their existing
order and semantics. No public API, dependency, ambient state, normalization,
compression policy or unsafe code is introduced. The proof is not cached on the
draft. See [source review](source-review.md) and [design proof](proof-design.md).

Ten new tests exercise actual commit hit/RDF fallback, equal-byte foreign Arc
owners, exact/over aggregate bounds, overflow, narrower limits, source-less
rejection, XML media classification, generated manifests, and exact noncompact
manifest/auxiliary XML preservation. Test-only hit telemetry adds no production
field, global counter or clock.

## Frozen matrix

Baseline revision is `c079b1c59`. Both executable pairs are bound to successful
release builds and complete source manifests. Rust 1.98.1, equal flags, CPU 2,
one worker, three warmups, thirty retained samples and A1/B1/B2/A2 order are
fixed in [protocol.json](protocol.json). The deterministic ordinary owned ODP
append lifecycle covers 64, 4,096 and 8,192 source slides, normal and allocator
instrumentation: 24 reports and 720 retained operations. Each report passes the
independent preservation, source/output identity and sequential-sink oracle.
Exact no-op, source mismatch and reversible patches are checked in corpus
preflight; source/candidate/sink identities are also checked per retained row.

| Repeat | Instrumentation | Shape | Baseline p50 ms | Candidate p50 ms | Delta | 95% median-delta interval |
|---|---|---|---:|---:|---:|---:|
| R1 | normal | tiny | 1.828 | 1.644 | -10.049% | [-10.365%, -9.690%] |
| R1 | normal | medium | 70.884 | 65.948 | -6.964% | [-7.228%, -6.689%] |
| R1 | normal | large | 142.287 | 131.806 | -7.366% | [-7.586%, -7.199%] |
| R1 | allocator | tiny | 1.989 | 1.804 | -9.295% | [-9.503%, -9.077%] |
| R1 | allocator | medium | 76.537 | 73.754 | -3.636% | [-3.810%, -3.411%] |
| R1 | allocator | large | 155.091 | 148.837 | -4.033% | [-4.166%, -3.804%] |
| R2 | allocator | large | 152.924 | 148.597 | -2.830% | [-2.976%, -2.577%] |
| R2 | allocator | medium | 77.084 | 74.290 | -3.624% | [-3.937%, -3.481%] |
| R2 | allocator | tiny | 1.990 | 1.814 | -8.836% | [-9.100%, -8.153%] |
| R2 | normal | large | 142.730 | 132.996 | -6.820% | [-7.133%, -6.540%] |
| R2 | normal | medium | 71.301 | 66.093 | -7.305% | [-7.439%, -7.129%] |
| R2 | normal | tiny | 1.815 | 1.652 | -8.968% | [-9.615%, -8.294%] |

All p95/p99, throughput and process-lifetime RSS observations remain in
[summary.json](summary.json). All twelve p50 confidence intervals are negative.
RSS deltas range from -2.030% to +1.708%; this is not a process-memory reduction
claim. Quantiles use midpoint p50 and nearest-rank p95/p99. Intervals use 10,000
seeded independent median-ratio resamples without correction for multiple
comparisons or run-order effects. No timing rerun or relaxed threshold was used.

| Allocator shape | Allocated-byte delta | Allocation-call delta | Reallocation delta | Deallocation delta | Regional peak delta | Retained-live delta |
|---|---:|---:|---:|---:|---:|---:|
| tiny | -2,087,682 | -1,039 | -93 | -946 | 0 | 0 |
| medium | -11,072,106 | -1,039 | -93 | -946 | -49,674 | 0 |
| large | -20,202,090 | -1,039 | -93 | -946 | -15,438 | 0 |

These p50 deltas match in both repeats. The allocated-byte reduction grows with
source size because the repeated final XML materialization/validation work is
avoided. The small peak reductions do not establish bounded-memory ordinary
snapshots or a flat whole-process memory bound. Copied/decompressed/recompressed
bytes were not separately instrumented here and no exact count is inferred.

## Supplementary diagnostics

Four separate large normal phase reports retain 120 operations. Commit p50
falls 12.686% / 13.080%; transaction falls 3.814% / 2.202%, snapshot opening
rises 0.981% / 0.474%, add rises 3.795% / 1.246%, and publication changes
-0.032% / +0.031% in R1/R2. The source change is in commit; unchanged-phase
variation is not attributed to the proof. Public-API phase clocks exclude
setup, checks and reporting and discard warmup rows. They support mechanism
review and do not replace the primary lifecycle gate. See
[phase-summary.json](phase-summary.json).

Two separate 100-operation whole-process perf runs include setup, warmups,
checks and reporting. Instructions fall 5.682%, cycles 6.424%, branches 5.352%,
branch misses 5.231% and cache misses 2.059%. Raw event runtimes and scaling
percentages are retained. These are not operation-only counters or exact causal
costs of the removed pass. Prior inclusive sampled profile shares overlap and
are retained as motivation, not current removable costs.

Retained assembly shows the candidate commit calling the proof predicate while
keeping the full validator available. The commit body grows from 3,045 to 3,395
instructions; the reported stack reservation remains 4,096 bytes. The validator
body remains 570 instructions and 536 bytes of stack reservation. The old
`to_bytes_bounded` target is absent/inlined in the candidate, which calls
`to_owned_package_bounded`; this does not mean serialization disappeared.
Generated-body counts are not executed instruction counts or peak stack/RSS.
See the raw assembly receipts and [manual review](assembly-review.md).

## Reproduction and final checks

[Source delta](source-delta.json) binds complete before/after text for both
changed Rust files to the compiled source epochs; `experiment.patch` preserves
the complete implementation and tests. The final owner suite passes 381 tests
with no failures or ignored tests. Warning-denied all-target Clippy passes.
Initial qualification/test-constructor compiler errors and the RDF test-fixture
failure remain archived with corrected passing retries; see
[integration notes](integration-notes.md).

All 387 harness tests pass (one ignored). Warning-denied rustdoc, scoped
formatting and crate boundaries pass. Live precleanup verification, fresh-copy
portable replay and resealed-summary tamper rejection pass. Cleanup removes
four staged executables totaling 233,058,712 bytes; the owned staging directory
is absent. The final seal covers 255 files (256 including the seal). Run `python3 -B verify.py --portable` in the
complete sealed bundle to authenticate capture/source/assembly identities and
recompute both summaries. Before cleanup, `--precleanup` additionally validates
retained executables and the final source epoch. `negative.py` requires a
resealed one-nanosecond summary mutation to fail. `finalize.py` records live and
fresh-copy replay, inventories the four staged executables and removes only
`/tmp/litchi-goal-0463`; `seal.py` refreshes the seal between proof steps.

This change adds no CRUD/source/output/corpus coverage. The registry remains
439 selectors / 36 defaults. No cold-cache, range-source, worker-scaling,
other-CRUD or Office GUI result is claimed. iWork is untouched. The
[next work](next-work.md) identifies broader source-backed semantic lifecycle
and producer/input/scaling evidence requirements.
