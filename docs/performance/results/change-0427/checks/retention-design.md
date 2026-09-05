# 0427 caller-drop retention design

This is a measurement design review for the allocator-only caller-drop
checkpoint. It is based on the current `tools/perf-baseline` allocator
observer and the four existing PPTX cross-copy lifecycle selectors. It records
scope and ownership decisions; it is not a result and authorizes no
optimization, cache-eviction, managed-budget, RSS, or latency claim.

## What the current evidence measures

The V3 allocator observer in
`tools/perf-baseline/src/allocation_metrics.rs` keeps absolute process-wide
callbacks and an observer-ordered region maximum. A lifecycle report's
`live_bytes_after` is sampled immediately after publication. In the owned
runner, the source and destination `Package`s, opened snapshots, plan,
published result, and sink are still in scope at that point. In the
source-backed runner, the source view, plan, publication result, caller
`Arc<InstrumentedSource>` values, and sink are still in scope; the editor is
consumed by `publish_cross_slide_copy_to_stream` and is dropped inside that
call. Reopen, semantic/raw/media checks, and teardown happen later.

Therefore the existing endpoint is an operation-exit observation with named
locals retained. It is not a post-caller-drop value. The existing
`region_peak_live_bytes` is likewise an operation-region maximum, not an
owner-attributed retained-size measurement. The absolute `live_bytes` value
includes the process baseline, runtime state, harness allocations, and any
other allocator callbacks observed in the region.

The 0423/0424 selectors remain the right matched workload: the owned and
source-backed plain cases exercise XML/plan state, while the media-rich cases
exercise the eight 2 MiB image leaves and their staged payload ownership. The
source-backed path uses an in-memory `ReadAt` observer, so this probe says
nothing about physical I/O.

## Checkpoint contract

The retention subcommand should build all corpus gates, expected output bytes,
hashes, and fixed row/report storage before the first retained sample. Per-
sample input clones and the bounded sink reservation belong inside the
whole-probe cycle, after the baseline checkpoint.
Each sample should use one fixed, preallocated row array and one allocator
region. No logging, hashing, JSON construction, dynamic vector growth, or
reopen should occur between checkpoints. Compare the emitted sink to the
already validated role-specific expected bytes with a byte comparison before
the first drop; report only the precomputed output digest. Finish the region
after the last drop and serialize its result afterward.

Use these rows in this order. Each row stores the raw absolute snapshot
(`allocation_calls`, `deallocation_calls`, `reallocation_calls`, failed calls,
allocated/deallocated bytes, `live_bytes`, lifetime high-water, overflow and
observer status) plus an explicit `held_owners` label.

| Row | Boundary and objects intentionally held | Interpretation |
| --- | --- | --- |
| `baseline_before_inputs` | Corpus, expected bytes, fixed row storage and report state; no per-sample inputs or sink | Process baseline for this sample. The corpus and harness are intentionally part of the baseline. |
| `prepared_inputs_and_sink` | Owned input `Vec`s, or source/destination caller `Arc`s, plus the reserved bounded sink | Setup ownership before document ingress. Input cloning and sink reservation are inside the whole-probe cycle but outside the named lifecycle operation. |
| `opened_documents` | Owned `Package`s and opened snapshots, or source-backed view/editor after public open | Catalog/document ownership after ingress and before planning. |
| `planned` | The cross-copy plan in addition to opened document handles | Plan staging and prepared closure are retained. |
| `published` | The returned publication result and sink, with document and plan handles still held | Exact output bytes have been compared; no semantic/reopen validation is performed here. |
| `drop_result` | Returned result dropped; all document, plan and sink handles still held | Records what the returned result releases. Its ownership must be observed; the probe must not assume that a result is metadata-only. |
| `drop_plan` | Plan dropped; document handles and sink remain | Isolates plan/patch/staged-payload ownership. For source-backed plans this also drops the plan's cloned source view; the caller source `Arc` remains. |
| `drop_document_handles` | Owned snapshots and packages dropped; source-backed view dropped. The source-backed editor is already consumed and dropped by publication. | Captures document/package release separately from caller-owned source release. |
| `drop_caller_source_arcs` | Source-backed caller `Arc<InstrumentedSource>` values dropped; not applicable to owned cases | Captures release of the caller's final source owners. It must not be described as cache eviction. |
| `drop_sink` | Reserved output sink dropped; expected bytes remain in the prevalidated corpus outside the sample | Final operation-local output-buffer release before the region ends. |

The implementation must not silently combine rows. In particular, the
source-backed `editor` has no post-publication local to drop: its by-value
publication API makes that destruction part of the `published` boundary. The
journal should record `editor_consumed_during_publish` rather than inventing a
separate editor-drop event. For the owned path, `opened::Snapshot` retains an
`Arc<OpcPackage>`; dropping snapshots before packages is required if the rows
are intended to distinguish those owners.

The region begins immediately before `baseline_before_inputs` and ends after
`drop_sink`. Emit its result as a nested `retention_probe` sample with a
distinct `retention_probe_region_peak_live_bytes` field. It covers setup,
publication, and drop callbacks and must never be substituted for the
historical lifecycle `operation_metrics.allocation.region_peak_live_bytes`.
The observer is non-reentrant and non-nestable, so checkpoint snapshots are
the phase evidence; a second allocator region cannot be opened inside the
first. Any unavailable, overflowed, underflowed, or observer-invalid row
fails the sample closed.

## Minimal capture

Use the current four lifecycle cases:

```text
pptx_cross_copy_plain_lifecycle
pptx_source_backed_cross_copy_plain_lifecycle
pptx_cross_copy_media_rich_lifecycle
pptx_source_backed_cross_copy_media_rich_lifecycle
```

The allocator binary is required. The new mode receives only these workload
flags (apart from the executable path and output location):

```text
retention --api owned|source-backed --corpus plain|media-rich
          --samples 30 --warmup 3
          --source-revision <40-lowercase-hex-digits> --output <report.json>
```

Run fresh child processes on CPU 2 with one worker for each case. Keep corpus
generation, all independent correctness/refusal gates, expected output
construction, and report-buffer allocation outside the sample loop. If a
before/after decision is later made, use the same four cases,
matched source/destination/output identities and the established fresh-process
`O/R1, S/R1, S/R2, O/R2` order for each role; do not compare normal elapsed
time with this allocator-only lane. A normal 100/10 guard may be collected
separately, but it is unnecessary for the drop-boundary question.

The header must bind the source revision and tracked source manifest, binary
hash and allocator identity, observer revision
`serialized_region_peak_v3`, the exact retention mode flags, CPU/worker
policy, case, corpus manifest, role-specific expected output hash and byte
length, and the verifier hash. The raw journal must preserve warmup/sample
counts and the exact held-owner labels for every row.

## Guardrails and interpretation

* `live_bytes` is an absolute process allocator count. Report checkpoint
  values and checked deltas, but do not call a delta “bytes owned by the
  plan” without an independent ownership proof.
* Allocator callbacks from any other process thread can enter the same global
  observer. Use one worker and fresh children, record the observer status, and
  fail closed on an invalid or incomplete sample. The observer mutex can alter
  scheduling; this is resource evidence, not timing evidence.
* `allocated_bytes` and `deallocated_bytes` are callback request totals.
  Reallocation overlap, allocator arenas, and physical RSS are outside this
  accounting. A return to the baseline live count does not prove that the
  allocator returned pages to the OS.
* A nonzero value after `drop_caller_source_arcs` can be runtime, allocator,
  cache, shared-owner, or harness state. It is not a leak diagnosis. For the
  source-backed path, record caller `Arc` strong counts when available, but do
  not treat them as a complete package-cache inventory.
* The output byte comparison must happen before drops and must use the
  prevalidated role-specific output. Source-backed and owned output bytes may
  differ lexically; cross-role byte equality is not a gate.
* Drop-window deallocation counts are useful descriptive evidence. They must
  remain separate from the operation allocation vector and from any latency
  claim. Do not subtract them from operation requests to manufacture a saving.
* The checkpoint does not exercise exact cache limits, one-byte-under refusal,
  pinned handles, oversized bypass, or managed execution-budget release. A
  full retention/cache conclusion still needs those near-limit rows and the
  phase-aware cache/RSS evidence listed in `retention-acceptance.md`.

The only safe immediate conclusion from a passing run is which absolute
process live-byte values were observed at the declared callback boundaries for
the named PPTX lifecycle. Any optimization conclusion requires the complete
matched source/output/oracle and provenance bundle plus an explicitly
predeclared interpretation of the drop rows.
