# 0729 — public DOC phase attribution with observer controls

The actual public DOC workflow gives a smaller finish fraction than 0728's
common-container control. In the profiled workflow, finish occupies
13.00–13.30% of NoHeadFoot.doc and 7.68–8.18% of FloatingPictures.doc. The small
fixture passes all 54 matched central control checks within 5%. The larger
fixture produces 17 control flags, so its instrumented fractions must not be
promoted to precise ordinary-path fractions or speedup predictions. Production
remains unchanged at `949a4a3037`; `performance_claim: none`.

| Fixture | Ordinary opaque p50, μs | Ordinary split p50, μs | Profiled empty p50, μs | Profiled clock p50, μs |
| --- | ---: | ---: | ---: | ---: |
| FloatingPictures.doc | 990.32–1135.86 | 991.65–1145.82 | 1001.26–1105.11 | 1021.09–1094.58 |
| NoHeadFoot.doc | 105.74–107.32 | 105.51–108.17 | 105.42–108.80 | 107.26–109.46 |

Each range spans nine independent process statistics. These are descriptive
ranges, not confidence intervals or pooled-sample estimates. The ordinary
opaque route times the complete public open/edit/replace/commit/output-copy
lifecycle. The ordinary split route adds separate outer clocks; profiled empty
uses the existing diagnostic APIs with empty callbacks; profiled clock adds
bounded semantic event timestamps. All routes retain output for untimed oracle
checks and include local commit/snapshot destruction in whole time.

| Fixture | Ordinary open % | Ordinary edit construction % | Ordinary replacement % | Ordinary commit % | Profiled finish % of whole |
| --- | ---: | ---: | ---: | ---: | ---: |
| FloatingPictures.doc | 25.95–29.22 | 7.72–8.83 | 19.97–22.45 | 37.72–42.82 | 7.68–8.18 |
| NoHeadFoot.doc | 22.17–22.51 | 7.69–7.88 | 32.15–32.78 | 36.40–36.77 | 13.00–13.30 |

These entries are ranges of per-process medians of same-owner sample ratios.
The ordinary columns come from ordinary split owners; the final column comes
from separate profiled-clock owners and must retain that qualification. The
full result reports output-copy and residual fractions, all semantic phases,
and p50/mean/p95/p99/maximum for every process. Residual time is not silently
assigned to a phase. Negative residuals are rejected because all five sequential
outer windows are contained in the monotonic whole window.

In profiled-clock owners, finish p50 ranges are 80.235–84.506 μs for the larger
fixture and 14.115–14.405 μs for the smaller. Public-reader validation occupies
17.17–18.29% at open and 17.20–18.13% at commit on the larger fixture; the smaller
fixture reports 12.33–12.45% and 11.36–11.62%. Those validation stages are
mandatory and are not nominated for removal. Even ideal elimination of the
small fixture's measured finish component would imply only roughly a 1.15×
ceiling for that instrumented scenario, before accounting for handoff
costs; it is not an achieved or predicted ordinary-path speedup.

Observer controls remain explicit:

| Fixture / comparison | p50 delta range | p50 flags / 9 | Mean delta range | Mean flags / 9 |
| --- | ---: | ---: | ---: | ---: |
| FloatingPictures, opaque → split | −9.75% to +15.70% | 4 | −6.13% to +14.20% | 5 |
| FloatingPictures, split → profiled empty | −9.24% to +7.58% | 4 | −7.89% to +2.40% | 2 |
| FloatingPictures, empty → clock | −5.63% to +5.44% | 2 | −3.45% to +4.85% | 0 |
| NoHeadFoot, opaque → split | −1.68% to +2.07% | 0 | −1.24% to +1.89% | 0 |
| NoHeadFoot, split → profiled empty | −2.35% to +2.90% | 0 | −2.06% to +2.62% | 0 |
| NoHeadFoot, empty → clock | −0.23% to +2.99% | 0 | −1.05% to +2.20% | 0 |

Positive deltas are slower. The prospective flag is an absolute change above
5% in either direction; it is an interpretation flag, not a retention gate.
All 108 central comparisons remain in the analysis. The 17 larger-fixture flags
comprise ten medians and seven means. They are not discarded as outliers or
attributed to a specific hardware/allocator cause by this experiment.

Source review identifies a real lifetime distinction in the diagnostic APIs:
ordinary open retains its strict editor through public-reader validation and
source retention, whereas the profiled strict-owner closure drops it earlier.
The empty-callback comparison therefore includes profiled implementation and
lifetime differences, not only callback dispatch. Every route uses the same
performance-diagnostics-enabled binary; default-build code generation is not
compared. Untimed trace/report construction also differs by route and can
affect allocator state between samples; the comparisons are not an isolated
per-callback instruction-cost measurement. The separate 16-event recorder calibration has process median ranges
of 400–440 ns across the two fixtures, but is not subtracted from workflow
measurements. Its full recorder state is kept observable. Only semantic events
are timestamped; no nested CFB attribution is claimed.

The fixed schedule contains three cycles and three rounds, rotating route order
and alternating case order. Both cases run all four routes in each round:
72 processes, three warmups per process, and 50 measured lifecycles per process
produce 3,600 measured owners and 216 warmups. The profiled-clock route contributes
900 owners and 14,400 checked semantic events. CPU 12, AMD EPYC 9R45, Rust 1.95.0,
release builds and warm OS caches match the current environment record.

The production source census is exactly 0728's. Both fixtures retain the same
45-UTF-16-unit paragraph-zero edit, output bytes, logical length proof, stream
inventory, metadata and direct semantic witnesses. The inherited 15 applicable
DOC case/control combinations reject all 540 executions across the matrix.
Qualification passes eight route processes, formatting, seven unit tests,
warning-denied Clippy and rustdoc. The initial argument-parser compile failure
and archived compatibility fields are preserved; neither is in the qualified
final probe. Review also corrected an initially mistaken suggestion to accept
negative residuals.

The analyzer and independently implemented auditor reproduce all process
statistics, event sequences/durations, same-owner fractions and matched
comparisons. Fifteen evidence-corruption controls pass. Exact post-cleanup replay
uses the executable identity witness; only the two owned build/binary roots are
removed. No allocation or RSS measurements are added, and the earlier allocation
record must not be treated as a measurement of these profiled lifetimes. No
cold-storage, remote-source, concurrent, cross-platform or native Office claim
follows. [Full analysis](results/change-0729/analysis.md).

The next bounded step is a source-qualified DOC batched validated-render handoff
pilot. The smaller fixture supplies a stable measured opportunity; the larger
fixture remains mandatory in ordinary-path retention and repeat controls rather
than receiving an observer-derived benefit promise. A candidate must retain
staging validation and final public semantic validation, prove exact no-ops,
successive edits and failure atomicity, bind rendering to the exact source and
policy, and account for retained-output memory and cloning. The existing
single-stream rendered handoff is a precedent, not proof that batched ownership
is safe. A generic unbounded render cache and a switch from Reuse to Rewrite
remain unsupported. Existing 0652 placement and 0663 preservation policy remain
binding. If the bounded lifetime proof fails, reject that design rather than
weakening the contract or silently changing this baseline.

[Evidence and replay index](results/change-0729/README.md). The broader non-iWork
goal remains active; no registered claim or CRUD coverage status is promoted.
