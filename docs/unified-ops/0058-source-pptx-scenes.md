# 0058 — Source-backed PPTX scenes and bounded retention

## Scope and provenance

Connect record 0057's retained shape projection to `SourceSlide::scene` and
integrate the L-PPTX-R1/R2 source-scene prerequisites from record 0010. The
weighted cache implementation, its reference-model tests and the tolerance
weight helper were adapted from candidate `b9587eebbb` as source-only changes.
No candidate history, binaries or measurement artifacts were merged.

The full goal remains active. Related layout/notes/comment/chart reads (R3),
shape property/text prerequisites (R6/R7), the unified PPTX adapter, all-format
writes, Python integration and paired performance acceptance remain open.
Existing borrowed shape and presentation APIs still need glossary migration.
No iWork source is changed.

## Source and ownership policy

The selected slide is validated through the existing `SlidePart` reader. Its
owned scene retains the actual `PartData` when XML is unchanged, including the
managed payload reservation; it does not call the source part's defensive
`blob_arc` copy. When MCE rewrites XML, only the processed buffer is retained.
Scene/shape handles contain no source presentation or slide owner. They remain
immutable, detached snapshots after their owner drops or converts to an editor.

Every call fences execution and exact source revision before lookup, after a
hit, after a miss/parse and after publication. A failed final publication fence
rolls back only that exact retained scene; an unrelated replacement cannot be
removed. Error paths also perform a final
fence. A deterministic versioned-source test changes revision after lookup and
verifies that the cached result is refused. A stale source cannot retrieve an
older cached scene because every future lookup repeats the revision check.

Declared physical part size and the existing semantic input-byte ceiling are
checked before materialization. Physical XML events, depth and attribute bytes
are checked over the original XML before MCE and semantic parsing. Ignored
extension/alternate branches therefore remain subject to caller limits.
Tokenization, tag-balance and duplicate-attribute violations are refused by
the physical scan with content-free text. The existing authoritative parser
remains responsible for document-root and namespace grammar.

This does not complete semantic execution-budget accounting: source payloads
retain their managed reservations, but parsed shape records, processed XML and
string arenas still use the existing shape reader's finite semantic limits,
not incremental execution-context charges. The broad semantic-policy migration
remains required; no full memory-budget conformance is claimed here.

## Cache policy and API

Each immutable presentation owns a weighted LRU with ceilings of 16 MiB and
4,096 entries. Clones share that memo. Keys are part names under the owner's
captured source generation; values contain only the independent scene.
Logical weights include XML retention, record/string capacities, tolerance
storage and modeled entry overhead. They are cache accounting, not allocator
measurements. The finite semantic memo is separate from the OPC payload cache.

Hits allocate no memo entry. Concurrent successful builders converge on the
first published scene. Oversized scenes and allocation-reservation failures
bypass retention; failed parses create no negative cache entry. Eviction may
release a cache share while a caller still pins the snapshot. Trimming releases
only unheld entries. Managed execution contexts disable scene retention
entirely so the memo does not keep a payload reservation alive for later calls.

`presentation::scene::CacheDiagnostics` contains only counters and gauges.
`SourceBackedPresentation::{scene_cache_diagnostics, trim_scene_cache}` expose
observation and clean trimming without publishing producer names or native IDs.
The new owned selector/handle policy from 0057 is unchanged.

## Verification

- Final scene/cache regression coverage: 18 tests passed, including a recency reference model,
  concurrent publication, weighted/count limits, oversized bypass, pin/trim,
  cold and warm source changes, post-hit fencing, managed cancellation and
  reservation release, editor conversion, source/eager parity, MCE and physical
  depth limits in a discarded branch.
- Glossary checker Python tests: 21 passed.

- `CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --all-features`: 893 unit,
  457 integration and 10 doctests passed (1,360 total).
- All-feature/all-target Clippy with `-D warnings`: passed after moving the
  test module below implementation items.
- All-feature rustdoc with warnings denied: passed.

- Workspace library compilation and both PPTX glossary profiles passed
  (11 bans per profile). Final workspace compilation, Clippy and rustdoc also
  passed after the publication fix.
- Review identified a post-publication fence race. Exact-scene rollback and
  `fence_discards` diagnostics now cover it. A deterministic test fails the last
  version check and verifies no entry/weight remains; a separate test proves
  rollback cannot discard a different scene under the same key.

All listed final gates pass. Changed Rust files were formatted;
`git diff --check` and the changed/untracked artifact audit passed. The three native SVG fixtures required by existing full-leaf
integration tests were temporarily copied from the base checkout, as in 0057.

## Commit and cleanup

No temporary worktree was created. `git add -A` failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only metadata filesystem. No commit is claimed; the stale partial index
must be refreshed before a later commit. Both required staged artifact checks
printed nothing. The local base tip remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`,
already an ancestor of HEAD.

The three temporary fixtures and only their newly created directories were
removed, along with generated lockfiles, Python cache and turn scratch.
`cargo clean` removed 29,935 files (30.1 GiB).
