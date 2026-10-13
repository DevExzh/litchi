# 0264 — PPT shape-text writes without whole-slide re-parses

Status: implemented on branch `uo2/ppt-scale`, cut from `feat/unified-ops-api`
at `a1120b3d3d`. It follows the performance finding of
[0260](0260-ppt-p0-surface.md) ("10k scaling: still super-linear") and the
coalescing and aggregate publication of [0218](0218-ppt-shape-coalescing.md)
and [0220](0220-ppt-aggregate-text-publication.md). Owned: `crates/litchi-ppt`.
The harness gained one profiling mode. The §4 bound change (owner approved
2026-10-14) raised two public constants; the glossary row "PPT chunked text
publication" was updated with it, and one `litchi-ops` test was adapted.

## Problem

Record 0260 measured the native 1k→10k same-slide text batch growing 35×,
and the unified batch 36×. A profile of the release harness (`perf record
--call-graph fp`, frame-pointer build, 10k native loop) put **60% of the
time in shape parsing**. Each native edit parsed the 10,000-shape slide
twice:

- **Target resolution.** The first edit hit the package's one-slide memo.
  Later edits started from a freshly opened committed package and parsed the
  slide again.
- **Commit readback.** `replace_shape_texts_batch` reopened the candidate
  and parsed the whole edited slide to check the targets' texts.

The 4096-operation hard bound needs three edits for 10k operations, so 10k
paid six whole-slide parses where 1k paid one. Single-shape sets showed the
same problem: `text_edit::Transaction::commit` re-resolved its target in the
published bytes through `resolve`, which parses every shape of the slide (66%
of a 1000-shape single set).

## What landed

### 1. Commit verification without a full re-parse

`src/text_edit/readback.rs` holds both verifiers.

**The guarantee of the reference verifier.** `readback_full` is the old
code, kept as is. It is used as the fallback and by tests. Its verdict for a
publication `P` with targets `T` is *accept* exactly when all of these hold:

1. The candidate opens as a presentation and lists its slides.
2. For every `t` in `T`, in order:
   - the slide at `t.slide` exists and has `t`'s persist identifier;
   - that slide's shapes parse. Structural errors are fatal. Content errors
     drop the shape and record a repair.
   - the shape at `t.shape` exists, and its text equals the replacement.

The single-shape commit checks the same properties through `resolve`. That
call also re-inspects the target's raw text atom, and it still does.

**The incremental verifier.** `readback_incremental` answers either *accept*
(with the shapes a parse would return) or *cannot prove*. It never refuses.
The production entry points are `readback` (batches) and `resolve_published`
(single shape). When the incremental verifier cannot prove a publication,
they run the reference verifier. So the production verdict and error equal
the reference verdict and error, provided that **an incremental accept
implies a full accept**. The incremental verifier accepts only when all of
these hold:

- the source slide's parse recorded no repair, so no shape was dropped and
  every position equals its `OfficeArt` index;
- the published slide's outline bodies equal the source slide's;
- a lock-step walk of the source and published `PPDrawing` payloads finds
  the same record tree:
  - the same record count, and the same kind, version and instance at every
    position;
  - byte-identical records everywhere except along the container chain of
    changed top-level shape containers. That chain is the enclosing
    `DgContainer`, the patriarch `SpgrContainer` and the `PPDrawing` length.
    Every published container must frame exactly its children.
  - inside each changed shape container, byte-identical children except
    exactly one `ClientTextbox`;
- each changed container belongs to exactly one target, matched by
  `OfficeArt` index (the walk replays the drawing parser's order, including
  background-shape filtering) and by the unchanged `Sp` shape identifier.
  Each target changed exactly one container.
- each changed container parses on its own. It converts at its index
  without being dropped and without a repair, keeps its identifier, and
  reads back its replacement.

**Equivalence argument.** `Slide::parse_shapes` is
`filter_some(convert(e_i, i, outline))` over the `OfficeArt` shapes `e_i` of
the drawing. The drawing-level checks (container framing, the record budget
and the nesting depth) depend only on the record tree, and the walk proved
that tree unchanged. Each `e_i` depends only on its own container's bytes
and its role. The role checks (`PATRIARCH`/`CHILD` against the topology) read
only the `Sp` atom, which is unchanged for every shape. The only changed
child, `ClientTextbox`, is opaque to the `OfficeArt` layer. Therefore every
unchanged shape converts to its source value. Each changed shape converts to
the value computed by the incremental verifier, at the same index with the
same outline. A full parse then returns the source array with the targets
replaced, records no repair, and holds each replacement at its target
position. The reference verifier accepts. Its checks of the slide position
and persist identity run on the same reopened slide list in both verifiers.

The walk is a structural check of the enclosing container chain. The
persist directory, `UserEdit` chain, `CurrentUser` and CFB checks remain the
reopen that both verifiers share. The published package's own open at
`Edit::commit` is unchanged.

**Differential tests** (`src/text_edit/readback_tests.rs`) publish through
the real rewrite and editor. A hook can corrupt the staged slide record
(simulating a faulty rewrite). Each case runs both verifiers on the
published candidate and requires three things:

- an incremental accept is a full accept;
- the proven shapes equal a full parse of the published slide (`Debug`
  form);
- for one-target publications, `resolve_published` returns exactly
  `resolve`'s result, or exactly its error.

| Test | Candidates | Full accepts | Incremental accepts |
|---|---:|---:|---:|
| genuine publications (1/3/40-shape decks; all, sparse and single targets) | 9 | 9 | 9 |
| every byte flip (3 masks) of the staged slide, two- and one-target | 4,707 | 2,999 | 926 |
| every header length (±1, ±2, ±8, −9) and kind/version/instance change | 484 | 113 | 73 |
| structural faults: sibling text, wrong or reverted target text, all `Sp` flag bits incl. background, identifier swaps, swapped and duplicated containers, truncation, extension | 89 | 62 | 5 |
| corpus (17 decks): every top-level shape the native edit accepts, genuine | 36 | 36 | 36 |
| corpus: byte flips across each of those publications | 5,526 | 4,655 | 522 |

No case had an incremental accept without a full accept. The single-commit
readback used its proof on 862 one-target candidates (374 flips, 452 corpus flips, 36 corpus genuine) and matched `resolve`
on all of them. The "full accepts only" cases exercise the fallback. For
example, a changed non-target sibling text is accepted, and only the full
verifier can prove it. Outline-backed placeholder text republishes the
document and is always verified in full.

### 2. Reusing the parsed slide across edits

The incremental verifier's accept yields the published slide's shapes as
`PatchedShapes`: the source `Arc` array plus the replaced positions. The
native transaction carries them only while its working artifact is that
exact allocation. Any other staged change invalidates them, checked with
`Arc::ptr_eq`. `Edit::commit` seeds them into the published package's
one-slide memo (the record-0260 memo) when both of these hold:

- the published bytes are the verified allocation;
- the source used default open policies, which the verifier's reopen also
  uses.

The memo materializes the array on first hit. A publication that is never
read again pays nothing. The next chunked edit, and the unified layer's path
resolution, hit the memo instead of parsing. Its own readback then starts
from that array. Invalidation is by construction:

- the memo is keyed by persist identifier within one immutable package;
- a seed is only ever placed into the package opened from the verified
  bytes.

Two tests pin this:

- `chunked_edits_seed_exactly_what_a_fresh_parse_returns`: three 50-operation
  edits;
- `sequential_single_edits_seed_exactly_what_a_fresh_parse_returns`.

Each checks that the memo holds the slide and equals a fresh parse, and that
the chained bytes equal edits on freshly opened packages.

### 3. Other profile findings

- **One buffer per drawing rewrite.** The grouped and single-target drawing
  rewrites copied every record into its own `Vec`, twice per container. They
  now append into one buffer and patch each container length in place. The
  per-record rebuild is kept in the tests as an oracle:
  `single_buffer_rewrites_match_the_per_record_rebuild` compares bytes,
  errors and match counts on every text shape of four decks and on
  corrupted drawings.
- **One open of the single-commit candidate.** The single-shape commit's
  validation open of the candidate also serves the incremental readback.
  It was opened twice, each time with a copy.

### 4. The 4096 hard bound: owner approved 2026-10-14, applied

With items 1 to 3 in place, an experimental build raised
`DEFAULT_MAX_OPERATIONS` and `MAX_SHAPE_TEXT_CHUNK_SELECTIONS` to 16384. It
was not committed; the diff is in the evidence JSON. The 10k batch then
publishes once:

| 10k batch | native ms | unified ms | per-op vs 1k (native / unified) |
|---|---:|---:|---:|
| 4096 bound (three publications), this record | 50.6 | 60.0 | 1.67× / 1.79× |
| 16384 bound (one publication), experiment | 32.5 | 35.5 | 1.07× / 1.07× |

Parsing is gone from the 10k profile: shape conversion is now only that of
the changed shapes, per operation. The remaining excess is per publication:

- each extra publication rewrites the whole 10,000-shape slide record;
- the append-only stream grows, so each edit writes a larger package (CFB
  write, unrelated-stream validation, reopen);
- the next edit materializes a patched shape array, and drops one.

Native phases at 10k: open + stage + commit = 2.1 + 50.7 + 2.7 ms with three
publications, against 0.4 + 32.9 + 1.1 ms with one.

**Policy proposal (approved by the owner 2026-10-14 and applied).** The
proposal: raise the primary PPT edit bound and the aggregate shape-text
selection bound from 4096 to 16384, and keep the 256-item inner chunk and
every other bound. Retained state per operation is the before/after text
of each change. The bound remains finite and is still checked before native
work.

What changed:

- `package::edit::DEFAULT_MAX_OPERATIONS` (and so `MAX_OPERATIONS`) and
  `slide_order::MAX_SHAPE_TEXT_CHUNK_SELECTIONS` are now 16384, with their
  doc comments.
- `MAX_SHAPE_TEXT_REPLACEMENTS` (256) and `MAX_SHAPE_TEXT_CHUNKS` (4096) are
  unchanged.

Effect on callers:

- `litchi-ops`' PPT coalescer, the query-fanout path and the harness derive
  their rollover from `MAX_OPERATIONS`, so they follow without code change.
- One past the native bound (16385 sets) exceeds the default unified
  operation bound (10,000), so `ppt_coalesced_set`'s rollover test now
  selects wider `SemanticLimits`.

Tests:

- `chunked_text_publication` now pins exactly 16384 selections as one
  publication, and 16640 selections as a refusal.
- The 4096/4097 chunk-count tests are unchanged.

Measured after the change: release, CPU 29, binary-level ABBA twice, 21
batch rounds per run, host load 7 to 8.5. Evidence:
[results/0264-ppt-bound.json](results/0264-ppt-bound.json).

| Batch | 4096 bound native / unified ms | 16384 bound native / unified ms | New / old (native / unified) | Pair ratio old → new |
|---|---:|---:|---:|---:|
| 1k | 3.061 / 3.349 | 3.031 / 3.335 | 0.990 / 0.996 | 1.096 → 1.099 |
| 4096 | 12.250 / 13.443 | 12.307 / 13.550 | 1.005 / 1.008 | 1.100 → 1.101 |
| 10k | 45.85 / 57.27 | 33.01 / 36.25 | 0.720 / 0.633 | 1.247 → 1.098 |

1k→10k is now 10.89× native and 10.87× unified: per-operation cost at 10k
is **1.09×** that of 1k on both legs (the target was ~1.2×). The 10k batch
publishes once. Native phases at 10k are open + stage + commit =
0.38 + 32.9 + 1.1 ms. The 1k and 4096 batches take the same path as before
and are unchanged within noise.

One consequence for the harness: the record-0220 routes
`ppt_text_batch_*_10000_aggregate_v2` derive their expected commit count
from the bound, so they now qualify one publication, not three. Their
output bytes therefore differ from the frozen 0220 evidence, which stays
unchanged. The routes were not renamed here (the harness change was
limited to new route names); the coordinator may want a `_v3` name.

## Performance

Harness: `tools/perf-baseline/examples/ppt_paired_probe.rs`. The method:

- release build, no LTO, pinned with `taskset -c 29`;
- binary-level ABBA (before, after, after, before) for every mode, and each
  run is itself ABBA native/unified;
- rounds per mode: single 61, cold 61, paths 41, batch 15, phases 9;
- opens, setup and byte-equality checks stay outside the timer.

The two binaries:

- **before** is the branch point `a1120b3d3d`;
- **after** is `9d51de790d`, this branch merged with trunk.

The host load average was 2.9 to 3.8. Evidence:
[results/0264-ppt-scale.json](results/0264-ppt-scale.json) holds the
pooled raw samples, medians, binary SHA-256s and the experiment diff.
Binaries are not committed.

| Scenario | Native before µs | Native after µs | After / before | Unified after / before | Pair ratio before → after |
|---|---:|---:|---:|---:|---:|
| single set, warm, 1 shape | 112.7 | 103.7 | 0.920 | 0.932 | 1.038 → 1.045 |
| single set, warm, 100 shapes | 279.9 | 134.7 | 0.481 | 0.496 | 1.018 → 1.045 |
| single set, warm, 1000 shapes | 1822.7 | 468.6 | 0.257 | 0.264 | 1.003 → 1.031 |
| single set, cold, 1 shape | 118.3 | 108.6 | 0.918 | 0.933 | 1.034 → 1.050 |
| single set, cold, 100 shapes | 431.5 | 299.9 | 0.695 | 0.691 | 1.011 → 1.001 |
| single set, cold, 1000 shapes | 3176.8 | 2029.5 | 0.639 | 0.628 | 1.003 → 0.991 |
| text batch 1k | 3293.6 | 3035.1 | 0.921 | 0.927 | 1.087 → 1.103 |
| text batch 4096 | 12176.8 | 12110.7 | 0.995 | 0.989 | 1.111 → 1.107 |
| text batch 10k | 108601.7 | 50626.5 | 0.466 | 0.490 | 1.130 → 1.189 |
| outline title set (SampleShow) | 316.5 | 311.4 | 0.984 | 0.976 | 1.040 → 1.048 |
| outline body set (41246-1) | 1377.9 | 1375.5 | 0.998 | 1.000 | 1.014 → 1.013 |
| paragraph set (basic_test) | 234.8 | 232.6 | 0.990 | 0.994 | 1.047 → 1.057 |
| notes set (basic_test) | 175.5 | 174.4 | 0.994 | 0.991 | 1.064 → 1.066 |
| slide add (SampleShow) | 248.0 | 245.3 | 0.989 | 0.991 | 1.037 → 1.036 |
| slide add (41246-1) | 1195.5 | 1193.1 | 0.998 | 0.999 | 1.009 → 1.009 |

No scenario regressed. Outline, paragraph, notes and slide-add sets do not
take the changed paths: outline text is verified in full, and notes and
slides are separate publishers. They are unchanged within noise.

Scaling, 1k→10k:

| Leg | Before | After | Per-op 10k / 1k before → after |
|---|---:|---:|---:|
| native | 33.0× | 16.7× | 3.30 → 1.67 |
| unified | 33.9× | 17.9× | 3.39 → 1.79 |

1k→4096 stays linear: 0.97× per operation after. The unified/native ratio
at 10k rose from 1.13 to 1.19 because native time halved; unified absolute
time also halved, and the ratio stays within the 2× batch budget. The
single-set pair ratios stay within the 10% budget. A unified 10k profile
(new mode `profile-batch-unified`) shows no slide parse in the unified
layer. Its remaining overhead is per run: patch and state drops, and native
edit open and commit.

## Unchanged guarantees

- Exact no-op bytes and signed refusal: the signed paths do not publish and
  are untouched.
- Encryption refusal, all-or-nothing batches, and exact durable and
  reversible patches: unchanged. Verification adds no state to patches, and
  the published bytes are byte-identical to before, as pinned by the rewrite
  oracle test and the harness's native/unified byte equality.
- The full verifier remains the fallback and the test oracle.

## Verification

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ppt -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --features ppt --all-targets -- -D warnings` | 0 |
| `cargo test -p litchi-ppt --all-features --no-fail-fast` (1,481 tests) | 0 |
| `cargo test -p litchi-ops --all-features --no-fail-fast` (1,878) | 0 |
| `cargo test -p litchi --all-features --no-fail-fast` (260) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ppt` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (21 profiles, 1043 bans) | 0 |

These gates were run after the §4 bound change, on trunk `f592aecf28`.
Trunk carries the fix for the DOCX query-digest failure seen at the
earlier merge (`4d1afe92bc`).

No capability or registry row changed, so no golden was re-blessed. Python
and the CLI do not expose the two constants and are unaffected.

## Follow-ups

- The remaining per-publication costs are O(file):
  - the CFB write of the growing append-only stream (`litchi-cfb`, not
    owned here);
  - the whole-slide record rewrite;
  - `Snapshot::from_bytes_with_options` at edit open.

  They bound multi-publication scaling even with the parse removed.
- The incremental verifier declines, and the full verifier runs, for:
  slides whose source parse recorded repairs; edits inside nested groups;
  non-`DgContainer` drawing roots; outline-backed text.
