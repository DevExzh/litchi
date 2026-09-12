# 0529 XLSX publication allocation instrumentation review

status: UNAPPLIED harness candidate; bounded source and patch review

This candidate adds operation-scoped allocator evidence for the publication
phase of `run_xlsx_cell_values_edit_save`. It is prepared against the exact
0529 base and has not changed the repository Rust source. OLE2/OOXML remains
the active priority; ODF is deferred and iWork is excluded.

## Identity and custody

| item | identity |
| --- | --- |
| baseline revision | `3f4c7be06159dc3d742d9c815800dc55bf7c2db6` (`3f4c7be06`) |
| baseline source | `tools/perf-baseline/src/lib.rs` |
| baseline source SHA-256 | `1d4f9ed81ffed0a65d3ff49572c50f2e8bf60867d4befeffc77d2174526fbafd` |
| baseline source Git blob | `3e7b26bf837e1916fd31864fd1e6f829fa039a04` |
| candidate copy | `/home/zhuhe/litchi-goal-0529-metrics/lib.rs` |
| candidate copy SHA-256 | `bdeeeef53b276cb72d813ff3c7345a302e72e0d8575e13d6757d07875189585f` |
| unapplied patch | `harness.patch` |
| patch SHA-256 | `ffa36478217bc16df797043e49147c5c487a5cbf01df97f8850e56e9a2dde5e2` |

The patch adds 83 lines to one harness source file and applies cleanly with
`git apply --check`. A disposable replay applied it to an exact baseline copy
and produced the candidate byte-for-byte. No hunk was applied to the working
tree during this review.

The candidate is bound to the current harness inputs:

| input | SHA-256 or identity |
| --- | --- |
| workspace `Cargo.toml` | `911a52cf6932b81550bc9ffd6e522c327dec178297ee9bc46d2ccedb693d4885` |
| `tools/perf-baseline/Cargo.toml` | `a04de024b9cbe9683cdb7307c3d7199daab7171bbdcb6bb8c30b361766857aca` |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |

## Exact candidate shape

`XlsxCellValuesSourceSummary` receives a
`publication_allocation_metrics: Vec<allocation_metrics::Sample>` field with
the same `skip_serializing_if = "Vec::is_empty"` envelope as
`commit_allocation_metrics`. Per-iteration evidence receives the matching
`Option<Sample>` field. `record_xlsx_cell_values` appends the publication
sample after the publication duration, preserving acquisition order and the
existing commit vector unchanged.

In `run_xlsx_cell_values_edit_save`, the commit region is finished before a
new publication region begins. The publication region starts immediately
before `publication_started`; the returned `MultiSnapshot` is held in the
existing `_published` block and therefore dropped while the region is active.
The duration is captured, converted to `publication_ns`, and only then is the
region finished. This keeps the returned snapshot's deallocations in the
publication sample while keeping the duration boundary unchanged. A disabled
normal binary publishes the explicit unavailable sample, matching commit
evidence. The allocator-metrics binary obtains the actual operation region;
its nonzero counters must be established by the separate allocator capture.

The two unrelated `XlsxCellValuesIterationEvidence` constructors (source
lifecycle and eager lifecycle) set the new field to `None`, so those evidence
records cannot fabricate publication allocator samples. No public library API,
dependency, unsafe code, commit measurement boundary, or existing allocation
region was changed.

## Focused harness test

`xlsx_source_cell_values_publication_allocation_metrics_align_with_phases`
checks three source-backed samples for vector length, phase-vector alignment,
the serialized field name and array shape, and the unavailable status/scope
used by the ordinary test binary. It also checks that the existing
`open + plan + commit + publication` duration identity remains exact. Cargo
lib tests, including those compiled with the allocator feature, do not install
the counting global allocator; consequently this test intentionally does not
claim measured nonzero counters. The allocator-only executable capture is the
authoritative check for those counters.

## Review disposition and required validation

Static review found the region ordering and ownership correct, and patch
replay was exact. Builds and tests were intentionally not run in this agent's
private-copy handoff, as required by the 0529 capture sequencing. Before
acceptance, the root coordinator should apply this patch before baseline
freeze, run normal and allocator builds, verify the publication vector in both
report forms, and confirm allocator samples are measured with nonzero
allocation activity in the allocator binary. The candidate remains unapplied
until that validation and the frozen OLE2/OOXML performance protocol complete.
