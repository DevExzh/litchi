# ODP shape attribute index experiment

The candidate is **rejected** under the frozen practical gate. Large normal
lifecycle p50 improves 2.660% and 2.628% in R1/R2, below the required 3% in both
repeats. Medium improves 3.465% and 3.192%; every normal confidence interval is
below zero. All regional allocator metrics are exactly unchanged, so there is
no practical heap benefit to justify the extra index state. No production
speedup is retained; 0460's accepted staging fusion remains the baseline.

## Mechanism and preservation

The measured candidate adds a private shape-only wrapper around `ElementAttrs`,
a typed enum for the 17 shape attribute queries, and a fixed first-occurrence
index. Common generic callers retain their old layout and generic scan.
Successful appends from lookup and harvest record the first recognized slot.
Values remain raw until requested; no decoded-value cache or heap map is added.

The state machine preserves exact namespace URI identity, aliases/rebinding,
unqualified/unknown-prefix handling, source order, duplicate and malformed
attribute reachability, and eager draw/presentation style fallback. Fresh decode
errors advance without appending; cached invalid values stay replayable. Harvest
keeps its separate error wording and appends only after success. Independent
raw-attribute and harvest oracles cover these cases. See [source review](source-review.md).

The layout diagnostic reports ElementAttrs 144 bytes, ShapeAttrs 424 bytes,
ShapeAttributeIndex 280 bytes, and ResolvedAttribute 80 bytes. The generic
getter's generated stack reservation remains 328 bytes; shape_builder grows
from 1,400 to 1,688 bytes, with a 360-byte get_known frame. Manual disassembly
shows direct indexed access on the normal known-key path, bypassing repeated
cached-prefix matching. A defensive fallback scan remains in the body. These
stack reservations are generated-code observations, not peak-stack or RSS
measurements.

The automatic `generated_code_evidence.known_cached_scan.eliminated` field is
not a valid proof: its direct-call heuristic returns a false positive even in
the baseline. [Assembly review](assembly-review.md) records this limitation;
the decision uses manual control-flow inspection and measurements. Raw outputs
and receipts remain unchanged and authenticated for both epochs.

## Frozen matrix

Baseline revision is `dbd2f8ece`. Both executable copies are bound to successful
release builds and complete source manifests. Rust 1.98.1, equal build flags,
CPU 2, one worker, three warmups, thirty retained samples, and A1/B1/B2/A2 order
are fixed in [protocol.json](protocol.json). The deterministic owned ordinary
ODP append lifecycle covers 64, 4,096, and 8,192 source slides, normal and
allocator instrumentation, with 24 reports and 720 retained operations. Every
report passes the independent preservation, source/output identity and sink
oracle. Exact no-op, source mismatch and reversible patches are checked during
corpus preflight; source/candidate/sink identities are also checked per row.

| Repeat | Instrumentation | Shape | Baseline p50 ms | Candidate p50 ms | Delta | 95% median-delta interval |
|---|---|---|---:|---:|---:|---:|
| R1 | normal | tiny | 1.826 | 1.783 | -2.362% | [-2.688%, -2.037%] |
| R1 | normal | medium | 70.765 | 68.313 | -3.465% | [-3.734%, -3.321%] |
| R1 | normal | large | 142.435 | 138.646 | -2.660% | [-2.913%, -2.225%] |
| R1 | allocator | tiny | 1.978 | 1.970 | -0.382% | [-0.845%, +0.117%] |
| R1 | allocator | medium | 76.289 | 75.096 | -1.565% | [-1.754%, -1.403%] |
| R1 | allocator | large | 155.796 | 152.430 | -2.161% | [-2.476%, -1.678%] |
| R2 | allocator | large | 153.714 | 151.660 | -1.336% | [-1.490%, -1.091%] |
| R2 | allocator | medium | 76.242 | 75.369 | -1.145% | [-1.659%, -0.882%] |
| R2 | allocator | tiny | 1.970 | 1.960 | -0.523% | [-1.283%, -0.242%] |
| R2 | normal | large | 141.966 | 138.235 | -2.628% | [-2.813%, -2.390%] |
| R2 | normal | medium | 71.065 | 68.797 | -3.192% | [-3.464%, -2.607%] |
| R2 | normal | tiny | 1.844 | 1.789 | -3.005% | [-3.380%, -2.549%] |

Every p95/p99, throughput and RSS result is retained in [summary.json](summary.json).
No elapsed or process-RSS result exceeds the +5% adverse threshold. RSS deltas
range from -4.737% to +1.876%, with no memory-reduction claim.
Allocation bytes, allocations, reallocations, deallocations, regional peak
above entry and retained-live deltas are exactly equal across matched lanes.
Allocator R1 tiny's confidence interval crosses zero; all normal intervals
remain negative. The practical gate still fails on both large normal rows.

Quantiles use midpoint p50 and nearest-rank p95/p99. Each interval uses 10,000
seeded independent median-ratio resamples, without correction for multiple
comparisons or run-order effects. No extra timing rerun or relaxed threshold
is used to change the outcome.

## Supplementary diagnostics

Four separate large normal phase reports retain 120 operations. Snapshot-open
p50 falls 3.948% / 5.682%, commit 3.663% / 4.171%, transaction
1.225% / 3.897%, and publication 0.186% / 0.207%; add rises
0.107% / 0.080% in R1/R2. These public-API clocks exclude setup, warmups,
checks and reporting. Unchanged transaction-phase variation is not attributed
to the shape index. Separate phase clocks do not override the primary gate.
See [phase-summary.json](phase-summary.json).

Two separate 100-operation whole-process perf runs include setup, warmups,
checks and reporting. Instructions fall 3.349%, cycles 3.014%, and branches
2.752%; branch misses rise 1.225% and cache misses 1.250%. Raw event runtime
and scaling percentages remain available. No operation-only counter claim
or causal attribution to every diagnostic shift follows.

## Reproduction and disposition

[Source delta](source-delta.json) authenticates the complete before/after text
of both changed Rust files against compiled source epochs; `experiment.patch`
retains the complete implementation and focused tests. Initial compilation
issues are preserved in the failed owner-Clippy receipt. The corrected retry
and 379 ODP tests pass. All 387 harness tests pass (one ignored). Both Rust
files are restored byte-exact to baseline. Final warning-denied all-target
Clippy, rustdoc, scoped formatting and crate boundaries pass. Precleanup,
fresh-copy portable replay and resealed-summary tamper rejection pass. Cleanup
removes four staged executables totaling 233,043,088 bytes; the owned staging
directory is absent. The final seal covers 255 files (256 including the seal).

Run `python3 -B verify.py --portable` in the complete sealed bundle to validate
capture/source/assembly identities and recompute both summaries. Before cleanup,
`--precleanup` also validates the retained binaries and final source epoch.
`negative.py` requires resealed summary tampering to fail. `finalize.py` records
precleanup and fresh-copy replay, then inventories and removes only
`/tmp/litchi-goal-0462`; `seal.py` refreshes checksums between proof steps.
[Integration notes](integration-notes.md) disclose the direct baseline-R1 capture
and later replay gate, initial compiler fixes, and source custody.

This experiment adds no CRUD/source/output/corpus coverage. The registry remains
439 selectors / 36 defaults. No cold-cache, range-source, worker-scaling,
other-CRUD or Office GUI result is claimed. iWork is untouched, and the full
non-iWork goal remains open. The [next work](next-work.md) shifts attention from
small matcher changes to commit proof reuse and broader semantic publication.
