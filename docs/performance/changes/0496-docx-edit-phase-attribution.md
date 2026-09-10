# 0496: DOCX opened-edit phase attribution

## Status and decision boundary

0496 is a diagnostic follow-up to the unresolved eight whole-child RSS flags
and three latency comparison cells retained by [0495](0495-docx-managed-document-edits.md).
It adds an opt-in phase field to the existing `docx-managed-edit` harness; it
does not change DOCX or OPC production code. The formal run now has 32 children
and 960 measured samples, with passing analysis, verification, and cleanup
receipts. The result is descriptive phase latency, whole-child RSS, and
full-lifecycle allocation evidence; it makes no CPU, speed, optimization, or
causal claim.

The purpose is to separate open, edit staging, commit, diagnostics, sequential
publication, published-snapshot drop, and commit drop inside the same opened
edit/save lifecycle. Publication is the largest named normal phase in all 16
normal children, but the managed short arm varies sharply between repeats. The
analysis retains the historical 0495 flags and records
`historical_flags_resolved: false`; it does not establish that those flags were
host noise or prove a cause.

## Diagnostic contract

`--phase-diagnostics` adds the optional `phase_diagnostics` row field. Without
the flag, the default serialized report remains unchanged. The diagnostic
contains these named intervals:

- `open_ns`
- `edit_staging_ns`
- `commit_ns`
- `diagnostics_xml_identity_ns`
- `publication_ns`
- `published_snapshot_drop_ns`
- `commit_drop_ns`

The intervals use wall-clock `Instant` measurements nested inside the
full-lifecycle clock. They are not CPU-time or CPU-cycle measurements. Checked
arithmetic records `phase_sum_ns` and a nonnegative `lifecycle_residual_ns`;
the residual covers phase-boundary arithmetic, budget evidence, result handling,
and package work not assigned to a named phase. `instrumentation_overhead_ns`
is explicitly `null` and unmeasured: phase-clock overhead is included in the
full lifecycle and is not isolated by a second control clock.

The allocator observer remains one full-lifecycle region around `execute_api`.
It is non-reentrant, so phase clocks do not create nested allocation regions or
phase allocation peaks. Full-lifecycle allocation counters and whole-child RSS
remain separate evidence. The package is consumed during publication, so this
diagnostic does not create a cache-after-publication or cache-after-drop gauge.

## Frozen comparison scope

The before side is production revision
`de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8`; the after side is
`44a4710699ef17041d5969240c30984dffbc3319`. Both use the same frozen
diagnostic harness source, while the before checkout also receives the retained
0495 harness overlay. Source manifests, patches, lockfile, build commands,
environment, and binary identities are retained in the evidence bundle.

The formal design has 32 fresh-child processes and 960 measured samples:

| Scope | Processes | Providers and roles |
| --- | ---: | --- |
| Before unmanaged | 12 | Owned, warm `FileSource`, and 4 KiB short-range; normal and allocator roles |
| After unmanaged | 12 | Owned, warm `FileSource`, and 4 KiB short-range; normal and allocator roles |
| After managed | 8 | Warm `FileSource` and 4 KiB short-range; normal and allocator roles |

Two reversed repeats use three warmups and 30 measured samples per process,
with CPU 2 and the shared measurement lock serializing capture. The 0495
corpus, exact output identity, semantic/media checks, source-version fence,
commit identity, replay/inverse/stale/foreign preflight oracles, source/sink
conservation, managed budget release, and allocator conservation remain in
the sample path. This scope adds no cold filesystem, native-producer,
borrowed-lifetime, real-network, concurrent-scaling, or atomic-save claim.

## Formal result and limits

The verified formal run has 32 children and 960 samples: 12 before-unmanaged,
12 after-unmanaged, and 8 after-managed children. Its paired unmanaged review
contains 294 comparison cells, of which 74 cross the threshold review gate.
The two reversed repeats use three warmups and 30 measured samples per child.
Managed-after rows are capability observations only; the paired before/after
review is limited to unmanaged rows.

Normal unmanaged lifecycle p50 latency is shown in milliseconds. Each repeat
keeps its before and after child paired; these are full-lifecycle values, not
phase values.

| Arm | Before R1 | After R1 | Before R2 | After R2 |
| --- | ---: | ---: | ---: | ---: |
| owned | 2.221 | 2.153 | 2.126 | 2.176 |
| file-warm | 4.794 | 2.302 | 2.279 | 2.326 |
| short | 5.359 | 5.340 | 5.238 | 5.419 |

Publication is the largest named phase by p50 in all 16 normal children. The
after-managed normal rows illustrate both the phase split and repeat variance:

| Arm | R1 lifecycle | R1 publication | R2 lifecycle | R2 publication |
| --- | ---: | ---: | ---: | ---: |
| file-warm | 3.563 | 2.465 | 3.553 | 2.453 |
| short | 6.509 | 5.406 | 4.092 | 2.993 |

The 74 retained threshold flags are separated by evidence scope:

| Flag family | Rows | Breakdown |
| --- | ---: | --- |
| Phase vectors | 60 | 18 commit-drop, 18 published-snapshot-drop, 10 diagnostics/XML identity, 7 open, 6 publication, and 1 phase-sum row |
| Full-lifecycle allocation | 9 | Reallocation calls p50/p95/p99 for each of the owned, file-warm, and short allocator arms; 185 to 228, +23.243% |
| Whole-child RSS | 4 | GNU-time child high-water RSS rows, including setup, preflight, verification, and report generation |
| Full-lifecycle latency | 1 | Short allocator p50: repeat 1 +61.39%, repeat 2 −44.02% |

The phase and lifecycle rows are not 74 independent production regressions.
The analysis retains repeat-level flags and descriptive paired bootstrap
blocks; `historical_flags_resolved` remains false, and no causal attribution
is authorized.

Full-lifecycle allocator p50 values are stable across the corresponding arms:

| API/phase | Allocation calls | Reallocation calls | Allocated bytes | Peak increment bytes |
| --- | ---: | ---: | ---: | ---: |
| Before unmanaged | 22,859 | 185 | 5,721,334 | 606,959 |
| After unmanaged | 9,696 | 228 | 1,623,696 | 609,903 |
| After managed | 20,535 | 328 | 4,270,196 | 622,454 |

Unmanaged before-to-after changes are −57.583% allocation calls, −71.620%
allocated bytes, +23.243% reallocation calls, and +0.485% peak increment.
Managed-after versus after-unmanaged changes are +111.788% calls, +162.992%
allocated bytes, +43.860% reallocations, and +2.058% peak increment. Every
allocator row deallocates the allocated bytes, returns live bytes to its entry
value, and reports zero failed allocation calls. These are full-lifecycle
allocator observations; phase clocks provide no nested allocation values.

Phase values are wall-clock `Instant` measurements, not CPU time. Publication
covers the 16,793,048-byte output copies into the preallocated retention sink;
`sha256_hex` and semantic/media verification follow `execute_api` and its
timing interval. These results therefore do not isolate production CPU or
establish a publication optimization. Whole-child RSS is separate and remains
unassignable to a named phase.

## Current custody and validation

The retained receipts cover the frozen source, capture, analysis, and build setup. The
diagnostic harness source is
`e3732d932356018dff0bd7a17b8eb38e60e177ee7d83992e9a04dbe82fe26c94`; before
and after normal/allocator build records are retained in
[`builds.json`](../results/change-0496/builds.json). The release library suite
reports 479 passed, 1 ignored, and 0 failed. Warning-denied Clippy, rustdoc,
crate-boundary, and rustfmt checks pass; the Python capture helper tests report
14 passed. [`seal-helper-tests-final2.json`](../results/change-0496/seal-helper-tests-final2.json)
reports 8/8 seal-helper tests passed; the documentation validators were rerun
separately. The interrupted debug attempt is retained under the development
subdirectory and is not an acceptance receipt.

The pre-capture source, build, environment, and ADR-refresh inputs are bound
by [`provenance.json`](../results/change-0496/provenance.json). The final
review is [`review-final.md`](../results/change-0496/review-final.md), the
formal analysis is
[`analysis/formal1.json`](../results/change-0496/analysis/formal1.json), and
the independent scalar cross-check is
[`scalar-review.json`](../results/change-0496/scalar-review.json). The formal
verification receipt passes all 32 children and 960 samples; `cleanup.json`
passes source-manifest and protected-file checks after removing 10.709 GiB of
disposable custody data and retaining four replay binaries. Cleanup and final
review are bound by the final seal inputs separately from the pre-capture
provenance. The live evidence seal remains a separate custody gate.

## Remaining goal boundaries

0496 adds no CRUD selector or production capability. Genuine borrowed input,
atomic filesystem save, independent producers, cold intersections, bounded
parallel scaling, durable history/composition, broad mutators, and the wider
non-iWork CRUD/security matrix remain open. The full non-iWork performance goal
remains open, and iWork is untouched.
