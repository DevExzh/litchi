# 0833 — bounded coverage-priority audit

Base: `eefaca16e39ace3c1e6219aedbac4fe1c0cd060a` (`perf(xlsx): keep first
column assignment inline`). This is a read-only prioritization note. It does
not report a new timing result or add iWork
evidence.

The selected 0833 qualification matrix is the following existing filesystem
coverage, each requested in `warm,cold-verified` state with the corpus and
oracles already owned by `tools/perf-baseline`:

- `opc_file_eager_open`
- `opc_file_source_open`
- `opc_file_eager_one_part_atomic_save`
- `opc_file_source_one_part_atomic_save`
- `pptx_file_eager_open_selected_slide_lifecycle`
- `pptx_file_source_open_selected_slide_lifecycle`

The OPC fixture is the fixed `FewLarge` incompressible corpus with four
4 MiB logical members ([`filesystem.rs:39-41`](../../../../tools/perf-baseline/src/filesystem.rs#L39)).
The PPTX lifecycle pair uses the existing 200-slide, eight-text-box,
eight-2 MiB-media corpus and checks source identity, selected-slide semantics,
and eager/source parity ([`README.md:1859-1885`](../../../../tools/perf-baseline/README.md#L1859)).
The six cases are a qualification boundary, not a before/after claim.

`cold-verified` is admitted only under the harness's strict regular-file,
page-aligned, allowlisted-filesystem, fsync/DONTNEED, fincore pre/post, and
positive `/proc/self/io read_bytes` rules ([`README.md:1770-1810`](../../../../tools/perf-baseline/README.md#L1770)).
It proves the recorded page-cache and process-read observations; it does not
prove physical-device temperature, device-cache state, or durable-storage
latency. The existing PPTX selected-slide lifecycle source replay proves
compressed-range locality, while eager records have no positional replay;
neither route by itself proves ordinary edit/save media passthrough.

## Ranked next work

### 1. Turn the selected six-case qualification into a current large-payload cold/warm baseline

After the qualification and offline reader checks, freeze a formal schedule for
the same six cases and report warm versus strict verified-cold results by route
and cache state. Keep the source/output hashes, save output bytes, logical
`ReadAt` vectors, materialization counters, selected-slide range coverage, and
every ineligibility reason separate. The OPC save pair is especially useful:
the harness seeds a destination and publishes through a same-filesystem
sibling plus atomic rename, while checking the per-state output hashes
([`README.md:1830-1857`](../../../../tools/perf-baseline/README.md#L1830)).

This is the highest-impact immediate measurement because the latest large OPC
publication record, [`0772`](../../0772-opc-mutated-save-current-baseline.md),
reports 59.403451 ms p50 for the four-large incompressible one-byte mutation
but explicitly excludes filesystem sync, cold storage, allocation, RSS, and
concurrency evidence ([`0772:17-31`](../../0772-opc-mutated-save-current-baseline.md#L17)).
The older ordinary-save record is limited to three small warm files and says
that its zero `read_bytes` observation is not physical-read proof
([`0819:17-40`](../../0819-real-file-ordinary-save-baseline.md#L17)); the
compaction trial likewise says identical compressed payloads show only
potential passthrough, not runtime copying ([`0827:214-225`](../../0827-ordinary-save-compaction-effect.md#L214)).

The result should answer only whether the current source/eager routes differ
for this larger untouched-member package and cache state. It should not be
turned into a physical-I/O or general Office claim. Any later optimization
must retain the save hashes, raw-member preservation, atomic-save behavior, and
the complete cold proof.

### 2. Measure targeted semantic updates on the existing media-rich DOCX/PPTX corpora, then attribute unchanged-media publication

Use the existing source-backed semantic controls after the filesystem baseline:
`docx_source_backed_one_edit_save`,
`pptx_source_backed_one_edit_save`, and the matched PPTX
`pptx_eager_batch_edit_save` / `pptx_source_backed_batch_edit_save` plus
single- and multi-slide batch controls. Their fixed corpora contain 200
paragraphs or slides and eight deterministic 2 MiB media members; the controls
already retain exact semantic, topology, relationship, media, raw unselected
member, patch, inverse, output-hash, and sink oracles
([`README.md:1888-1978`](../../../../tools/perf-baseline/README.md#L1888)).

The measurement should separate open/planning, commit, and publication phases
and report whether untouched media remains source-backed through publication,
whether compressed members are copied without logical decompression/copying,
and which edits expand the dependency closure. The existing cross-copy results
show why this is worth measuring: eight 2 MiB images reduced media-rich
cross-copy lifecycle p50 from 410.081 ms to 183.024 ms when verified
source-compressed bytes were transferred ([`0742:17-28`](../../0742-pptx-owned-cross-copy-media-transfer.md#L17));
those results are cross-copy evidence, not ordinary edit/save evidence. The
follow-up digest-reuse result also reports 33.56 MB fewer allocations and
16.4 MB lower peak live bytes in that cross-copy path
([`0751:23-55`](../../0751-pptx-cross-copy-apply-digest-reuse.md#L23)).

The next implementation question, if the measurements justify one, is the
existing preservation-sensitive path from [`GOAL.md:441-456`](../../../GOAL.md#L441):
copy untouched compressed ZIP payloads and large binary media directly while
reserializing dirty members, with complete framing, ZIP64, descriptor, sink,
and fallback checks. Do not claim that the selected PPTX lifecycle pair has
already established this behavior; its timed scope is open plus selected-slide
access, not semantic publication. Likewise, the current XLSX inline-map result
is intentionally bounded to its pinned real edit and promotion guard: it says
there is no general nonempty-column latency or broad cold/range/concurrency
claim ([`0832:44-62`](../../0832-xlsx-inline-column-map.md#L44)).

### 3. Close the explicit range and bounded-concurrency gaps with the existing deterministic controls

Run the OPC range-source pair on the same `few-large` incompressible shape using
the existing fixed-latency, bandwidth, request-overhead, and maximum-range
simulator: `opc_range_source_open` and `opc_range_source_open_main_read`.
Retain the exact logical/physical request counts, request-size buckets,
compressed-range overlap, materialization count, and payload hash. The harness
documents this simulator as caller-source evidence rather than network or disk
measurement ([`README.md:2630-2643`](../../../../tools/perf-baseline/README.md#L2630)).

Then run the already-bounded OPC cache matrix over its fixed many-small
incompressible corpus: `opc_source_cache_budget_boundary`,
`opc_source_cache_control_contention`, and
`opc_source_cache_managed_contention`, at worker widths
`1,2,4,8,available`, retaining lock diagnostics only when explicitly selected
([`README.md:1675-1705`](../../../../tools/perf-baseline/README.md#L1675)).
This gives a current source/cache baseline for flights, waiters, retained
bytes, budget refusal, worker width, and request throughput before changing
single-flight or scheduling behavior.

The prior scaling evidence is useful but stale for this purpose. `0786` used
immutable in-memory sources and explicitly did not establish physical cold,
delayed range, native CRUD, or cross-session contention
([`0786:17-21`](../../0786-finite-execution-budget-scaling.md#L17)); `0816`
reported 7.364x delayed CFB and 5.674x delayed Part width-eight speedups but
also identifies the source as a synthetic delay model rather than a real
network or physical-cold measurement ([`0816:1-16`](../../0816-delayed-source-budget-scaling.md#L1)).
The new run should therefore remain a descriptive transport/scheduling result,
not a production-wide speedup claim. Any scheduler or cache change still needs
the explicit worker, memory, I/O, and scaling evidence required by
[`GOAL.md:568-606`](../../../GOAL.md#L568) and the checklist's measured-claim
gate ([`CRUD_Scenario_Checklist.md:71-81`](../../../CRUD_Scenario_Checklist.md#L71)).

These priorities follow the definition of done: targeted updates should avoid
unrelated parsing/recompression, unchanged large media should flow from source
to output without unnecessary logical copies, and bounded parallel paths must
show real scaling ([`GOAL.md:799-817`](../../../GOAL.md#L799)).
