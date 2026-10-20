# 0357a — record-sized scoped PPT family patches

- Status: implemented on `uo2/w3-ppt` at `e673d4addb` (follow-up of record
  [0357](0357-ppt-api-contract.md) §9).
- Date: 2026-10-20.
- Input: the coordinator's requirement that a scoped PPT family patch be
  record-sized (what changed), not whole-package bytes, while keeping exact
  apply and inverse semantics.
- Owned: `crates/litchi-ppt/src/scoped.rs`, `crates/litchi-ppt/src/package/owner.rs`
  (a digest cache), the family tests that read the removed accessors, and
  the probe `crates/litchi-ppt/examples/scoped_patch_probe.rs`.
- `performance_claim`: measured (§4).

## 1. What changed

`scoped::BytePatch`, the representation behind every `family!` patch (the 12
families of record 0357 §2: `font`, `broadcast`, `picture_bullets`,
`animation`, `document_comparison`, `external_media`, `diagram`,
`slide_sync`, `client_data`, `master::metadata`, `master_layout`,
`embedded::object`), held the complete source and target package bytes. It
now holds:

- the source and the target endpoint, each as its length and SHA-256
  (`litchi_core::patch::BlobId`, the helper behind
  `slide_order::artifact_hash`);
- the changed ranges: the two images compared in 512-byte blocks (one OLE
  sector), adjacent changed blocks coalesced into one range
  `(offset, before, after)`. Bytes past the shorter image belong to the last
  range, so a range's two sides differ in length only there. A no-op patch
  has no ranges.

The ranges are shared (`Arc`) between a patch, its clones and its inverses;
`inverse` swaps the endpoints and flips a direction flag instead of copying
bytes. Equality compares the endpoints and the oriented ranges, so
`patch.inverse()` equals the diff made in the other direction.

Semantics, unchanged in effect:

- **Commit.** `Base::publish` and `Base::publish_owner` produce the target
  bytes exactly as before. The commit computes the ranges (one linear
  compare, which also decides the no-op) and the target's SHA-256. A no-op
  publishes the source package unchanged, as before. Its patch now carries
  the source digest, so the first no-op commit of an unhashed package hashes
  it once.
- **Apply.** `apply_x` captures the package's exact bytes and refuses it
  with `PatchConflict` unless its length and SHA-256 equal the patch source.
  It then copies the bytes, resizes them to the target length, writes each
  range's target side, and reopens the result with the source package's
  options. The source is bound by length and SHA-256, so the rebuilt bytes
  are the exact target by construction. Debug builds hash them again and
  assert the target digest; release builds skip that redundant hash (the
  patch has no public constructor, so no caller can forge one). A changed
  apply on a signed package is still `Signed`.
- **Inverse.** It is the same operation in the other direction: on the exact
  target it rebuilds the exact source, and on anything else it is
  `PatchConflict`.

Digest cost: `Package` caches the SHA-256 of its exact bytes in a
`OnceLock`, computed on first use. The scoped publisher seeds the cache of
every package it publishes with the target digest it already computed. A
commit therefore hashes one image, the target. The source is hashed only the
first time a package is edited or receives a patch, and never again when it
came from a scoped commit. Applying a patch to the package of the commit
that made it, or of the replay of that commit, hashes nothing.

## 2. Public surface

Every family `x::Patch` loses `before_bytes()` and `after_bytes()`, which
exposed whole artifacts. It gains three content-free accessors:
`source_len()`, `target_len()` and `changed_bytes()`. `changed_bytes()` is
the bytes the ranges span, counting the longer side of each range, and is
zero for a no-op. `Debug` prints `source_len`, `target_len`, `ranges`,
`changed_bytes` and `payload_bytes` (both sides). No consumer outside the
crate read the removed accessors. Inside the crate, the eight tests that
asserted `patch().before_bytes() == source` now assert `source_len` and
`target_len`.

`package::metadata::Patch` and `package::edit::Patch` are not `family!`
patches and still carry whole-artifact bytes. `package::edit::Patch` wraps
the slide-order durable patch, and its consumers in litchi-ops size
snapshots from those bytes. Both are left for a separate record.

## 3. Tests

- `scoped::tests` (4 tests): a no-op diff has no ranges and refuses a
  package that differs by one byte or is one byte shorter. A mid change
  spans exactly its coalesced blocks. Growth, shrinkage, a shrink to a block
  boundary, and the empty image all put the length change in the last range.
  A same-length package that differs outside the ranges is refused. Every
  case checks apply, inverse, a double inverse, equality with the reverse
  diff, and refusal of the wrong source.
- `tests/scoped_families.rs` `external_media_edits_publish_replay_and_invert_exactly`
  now also asserts that a one-path change is record-sized:
  `changed_bytes()` is at least the growth and at most the growth plus 8
  sectors, and less than half the source. On `test-data/unified-ops/0294/video.ppt`
  (43,520 bytes) the patch spans 13,312 bytes: the appended 11,264-byte
  tail (22 new sectors carrying the appended records) plus four existing
  sectors that the publisher rewrites in place. The CFB header is unchanged.
  The inverse spans the same bytes, and its endpoints are swapped.

## 4. Performance

Probe `scoped_patch_probe` on `test-data/unified-ops/0294/video.ppt`: 1,000
independent single-operation `external_media` commits on one source
(`set_path` of movie 1 to a distinct path each time). Every patch is kept,
as a patch log or undo history would keep it, and the packages are dropped.
The probe then applies each patch to the source and each inverse to that
result, and checks the restored bytes outside the timed regions. A chained
history (each commit editing the previous result) was not measured. The
append-only publisher grows the package about 11 KiB per commit, so 1,000
whole-artifact patches would retain about 11 GB on this host, which has no
swap.

Method: release builds with no LTO, `taskset -c 29`, binary-level ABBA
(before, after, after, before) × 10 rounds, so 20 runs per build. Peak RSS
comes from `/usr/bin/time -v` (`%M`). The times are pooled medians of each
run's per-operation median. "Before" is `e673d4addb`, built from
`git archive` into a separate target directory with the same `Cargo.lock`
and probe source. The binaries are not committed; their SHA-256 are in
[`results/0357a/binaries.sha256`](results/0357a/binaries.sha256). The raw
runs are in [`results/0357a/abba.tsv`](results/0357a/abba.tsv), the summary
in [`results/0357a/summary.json`](results/0357a/summary.json), and the
script in [`results/0357a/abba.sh`](results/0357a/abba.sh) and
[`summarize.py`](results/0357a/summarize.py). Load average was 4.6–5.6
(other agents' builds on other CPUs).

| Metric | Before | After | After/before |
|---|---:|---:|---:|
| bytes retained per patch | 98,304 | 15,360 | 0.156 |
| bytes retained by 1,000 patches | 98,304,000 | 15,360,000 | 0.156 |
| peak RSS (KiB, median; range) | 110,910 (110,816–111,264) | 19,758 (19,584–19,900) | 0.178 |
| commit, median µs | 195.9 | 217.0 | 1.107 |
| 1,000 commits, total ms | 200.2 | 224.4 | 1.121 |
| apply on source, median µs | 23.0 | 20.5 | 0.892 |
| inverse apply on target, median µs | 19.1 | 18.4 | 0.962 |

The retained bytes per patch are both sides of every range: 2,048 source
bytes (four sectors) plus 13,312 target bytes. Before, a patch held both
whole images (43,520 + 54,784 bytes).

The commit costs about 21 µs more (+11%). That is one SHA-256 of the
54,784-byte target (about 2.5 GB/s on this host), the minimum an exact
inverse needs: without a digest of the whole target, a package that equals
the target only inside the ranges could not be told apart from the target.
The coordinator's budget allowed one compare and two digests per commit.
The source digest is cached, and the commit no longer copies the target
into the patch. Apply and inverse apply got slightly faster: they compare
32-byte digests instead of whole images, and the digest of a published
package is already cached.

## 5. Gates

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ppt --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ppt --all-targets --no-default-features -- -D warnings` | 0 |
| `cargo test -p litchi-ppt --all-features` (1,525) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ppt` | 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py --profile ppt-minimal` / `ppt-full` | 0 / 0 |
| `cargo check -p litchi-ops --all-features --all-targets` | 0 |

## 6. Follow-up: closing the 0357 §9 family gaps

Four commits on `uo2/w3-ppt` after `ac9e036031` close the animation, object
and master gaps that [0357 §9](0357-ppt-api-contract.md) listed.

- **Animation net change** (`6103aab069`). `animation::Edit::commit`
  compares each staged slide or main-master record with its source. It no
  longer tracks whether a verb rewrote a record. A record whose bytes equal
  the source, or whose animation model equals the source model, is unstaged
  and keeps its source bytes. The model is the timeline, the build list and
  the legacy shape animations, the last compared by serialized
  `AnimationInfo`. A change reverted before the commit is therefore an exact
  no-op. A partial revert publishes only the records that still differ.
- **Signed decks open** (`6103aab069`, `17a5751052`). The record publisher
  `embedded::object::Editor` has an open that admits a digital-signature
  component. It records the signature, and `finish` refuses any changed
  publication. Encryption and rights-management components are still
  refused. `edit_animations()` and `edit_objects()` now open a signed deck.
  An exact no-op commit keeps the signed bytes, and a changed commit is
  `Error::Signed`, as in every other family.
- **Objects on a deck without a list** (`17a5751052`). `edit_objects()`
  edits an empty list (seed 1) when the live document has no
  `ExObjListContainer`. The first committed change inserts the container
  directly after the `DocumentAtom`, which is where [MS-PPT] 2.4.1 places
  `exObjList`. A list that already exists is replaced in place. A fix on
  the way: removing an object whose storage the same edit had staged is now
  a no-op.
- **Notes masters and incremental saves** (`de984025c7`). [MS-PPT] 2.5.12
  requires a null `NotesAtom.slideIdRef` in the notes master. PowerPoint
  writes 0x80000000 there, and other producers write 0x80000001, often with
  `fMasterScheme` set. Both values are in the `MasterId` domain (2.2.16),
  never `SlideId` values. The two master validators and the master-layout
  inventory now accept two forms: a null reference with null defined flags,
  or a reference in the `MasterId` domain. A slide reference is still
  refused. `Presentation::masters()` now reads the live `DocumentContainer`,
  so it no longer refuses an incrementally saved deck. Before the fix,
  `masters()` failed on 14 of the 30 corpus decks that open, and both master
  families refused the notes master of 13 of them. Now every deck lists its
  masters, and every listed master opens both families. Title-master
  coverage uses Apache POI `slide_master.ppt`, copied into `test-data/poi`
  (Apache-2.0; provenance in `test-data/poi/PROVENANCE.md`).
- **Signed-package tests** (`348bfe5e65`). Every family now has a signed
  test (admit, no-op, `Signed` on a changed commit), including
  `master::metadata` and `master_layout` on main, handout, notes and title
  masters.

Gates on `348bfe5e65`: `cargo fmt --all --check` 0; clippy
`--all-features` and `--no-default-features` (`-D warnings`) 0 / 0;
`cargo test -p litchi-ppt --all-features` 0 (1,540 tests); `cargo doc` with
`-Dwarnings` 0; `check_api_glossary.py` `ppt-minimal` / `ppt-full` 0 / 0;
`cargo check -p litchi-ops --all-features --all-targets` 0; the litchi-ops
PPT tests (`crud_matrix`, `ppt_unified_read`, `ppt_unified_write`,
`ppt_elements`, `--lib`) 0 (988 tests, no re-bless).

## 7. Follow-up: restoring the routes record 0357 made crate-private

Record 0357 kept four native capabilities crate-private with no public route
([0357 §9](0357-ppt-api-contract.md#9-gaps-and-follow-ups)). The owner's
direction is no capability loss. Three commits restore them in the contract
shape. The scoped families' patch representation (§1) is unchanged.

### 7.1 Durable patches, joins and three-way plans (`910d5de31d`)

These live on the primary family `package::edit`, not on a parallel owner:

- `package::edit::Patch::to_durable(&self, PatchLimits) -> Result<litchi_core::patch::Patch<Reversible>>`
  and `Package::apply_durable<Mode>(&self, &litchi_core::patch::Patch<Mode>) -> Result<Commit>`.
  Both delegate to the native `Patch::to_durable` / `Snapshot::apply_durable`.
  The commit publishes a `Package` with `Diagnostics`, and
  `operation_count()` is the number of durable operations. The replay's own
  patch is exact, so `apply` and `inverse` work on it. It keeps no semantic
  operations, so `to_durable` on it is `Refusal::UnsupportedPatch`. Convert
  the durable patch itself instead. A bound the limits cannot hold is the new
  `Error::Durable(PatchError)`, which litchi-ops classifies through its
  `PatchError` rule.
- Undo/redo is `patch.inverse()` plus `Package::apply`. The native
  `History` alias and `Snapshot::history` are deleted.
- `Package::plan_three_way(&self, &Patch, &Patch) -> Result<ThreeWayPlan>`,
  with `package::edit::{ThreeWayPlan, Conflict}` (`is_clean`, `conflicts`,
  `Conflict::target`). A patch from another package is `PatchConflict`. The
  native plan has no apply or merge, so no apply or merge verb was added.
- The native composition (`prepare_move`, `compose`, `PreparedMove`,
  `Composition`, `JoinError`) joined independently prepared slide moves over
  one source whose position ranges were disjoint. It is now
  `package::edit::Patch::join(&self, &Patch) -> Result<Patch>`, using the
  glossary's canonical `join`. It joins two move-only patches over the same
  source whose permuted position sets are disjoint. Such moves commute. The
  joined patch replays both patches' moves in one native commit, and its
  operation count is the sum of the two. Another source is `PatchConflict`.
  Other operations, or overlapping moves, are the new
  `slide_order::Refusal::UnsupportedPatch` (kind `UnsupportedOperation`). This
  is more general than the old composition, which took single moves only. The
  native types are deleted.

Tests: `slide_order::durable_tests` moved to `tests/durable_patches.rs` on
the public route (11 tests). These cover insertion, notes, comment and
outline-text durable replay and revert, the durable bound, the empty durable
patch, undo/redo, joins (commuting, durable round trip, overlap, other
operations, other source) and three-way plans. They include the
native-handoff visual-transition scenario (exact replay and inverse, durable
JSON round trip, undo/redo, three-way conflict). The in-crate history and
composition unit tests use inverse/apply and the native join.

### 7.2 The source-backed text splice family (`0f78d32842`)

`litchi_ppt::text_edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`
is registered in `ppt-minimal` and `ppt-full`, with `is_changed` banned in its
scope:

- `Package::edit_text(&self, text_edit::Target) -> text_edit::Result<text_edit::Edit>`;
  `Edit::{text, target, set_text, commit}`;
- `Package::apply_text(&self, &text_edit::Patch) -> text_edit::Result<text_edit::Commit>`;
- `Commit::{package, into_package, patch, diagnostics, into_parts}`;
  `Patch::{is_noop, inverse, target, source_fingerprint, target_fingerprint}`;
  `Diagnostics::{changed, target, changed_spans, replacement_bytes,
  source_bytes, source_fingerprint, target_fingerprint}`; `LimitKind` has no
  caller-selected dimension.

The edit opens `SourceSnapshot::from_shared` over the package's retained
compound-file index and resolves the atom with bounded positional reads. The
commit is the native owner, unchanged: one `ValidatedOverlayPlan` splice in
`PowerPoint Document`, the composed target reopened, and the text read back.
The commit then adopts the candidate `SharedOleFile`, which the native owner
already opened and validated over the composed source, as the published
`Package` (`SourceBackedPackage::from_validated_shared` plus
`Package::from_source`). It does not index the container a third time, and it
never materializes the package. The published package's `write_to` and `save`
stream from the composed source. A partial sink receives exactly the accepted
prefix.

The patch is the record-sized `SourceBackedPatch`: the atom offset and its two
byte images, bound to the exact source and target artifact fingerprints.
`apply_text` accepts any package whose fingerprint is the patch source, so a
reopened, byte-equal package also qualifies. It refuses every other package
with the new `text_edit::Error::PatchConflict`. A signed package opens. Its
text is read with the read-only resolver, and an exact no-op commit returns it
unchanged. A changed commit is the new `text_edit::Error::Signed`. The two
variants are classified by litchi-ops (`invalid_input`,
`document_protected`), and the scoped and slide-level mappers handle them.
`text_edit::SourceSnapshot` stays as the source-backed read view. The
crate-private whole-artifact text owner's `Commit` and `Patch` are renamed
`RecordCommit` and `RecordPatch`.

Tests (`text_edit::family::tests`, 7): commit, replay and inverse on eager
and positional packages, with the unrelated stream kept and the length
unchanged. They also cover replay on a reopened equal package, conflicts,
chained splices reading through the composed source, no-ops, length and
absent-target refusals, the macro refusal, the signed rule, and a partial
sink. The native `SourceSnapshot` and `SourceTransaction` unit tests stay.

**Measurements.** Probe `crates/litchi-ppt/examples/text_splice_probe.rs` on a
Builder-made 1,000-shape deck (10 slides of 100 text boxes, 168,448 bytes).
Each iteration opens the deck from its file (`Package::open`), makes one
same-length replacement of slide 5 / shape 50, commits, and streams the
published package with `write_to` to a counting sink. There are 50 iterations
per run. The `text` mode is `edit_text`. The `primary` mode is
`Package::edit().set_shape_text`, the only public route at the parent commit.

Method: release, no LTO, `taskset -c 29`, binary-level ABBA × 10 rounds.
"Before" is `910d5de31d`, built from `git archive` with the same `Cargo.lock`
and the probe with its `text` arm removed. "After" is `0f78d32842`. Every slot
runs `primary`, and the after slots also run `text`. Load average was 5.4–11.4
(other agents' builds). The scripts and evidence are in
[`results/0357a/text/`](results/0357a/text/): `abba.sh`, `range_abba.sh`,
`summarize.py`, `abba.tsv`, `range_abba.tsv`, `summary.json` and
`binaries.sha256` (probe, perf-baseline and deck SHA-256; no binaries are
committed). Times are pooled medians of per-run medians.

| Route (1,000-shape deck) | open µs | edit+commit µs | commit µs | write_to µs | total µs | written bytes | peak RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|
| before, `primary` (20 runs) | 27.3 | 549.5 | 40.6 | 2.2 | 579.1 | 184,320 | 6,074 |
| after, `primary` (20 runs) | 27.2 | 551.0 | 40.3 | 2.2 | 579.7 | 184,320 | 6,242 |
| after, `text` (20 runs) | 22.8 | 827.8 | 520.4 | 5.2 | 856.1 | 168,448 | 5,010 |

- The primary edit is unchanged: total 1.001×. Its peak RSS medians differ by
  2.8%, and the two ranges overlap (5,856–6,316 against 6,076–6,432 KiB).
- The restored splice costs 1.48× the primary edit's time on this small
  in-memory-sized deck. It keeps the file length (168,448 bytes against
  184,320; the append-only primary publisher adds a user edit), uses 20% less
  peak RSS, and never holds the package bytes. Its commit is the native
  owner's validated publication: an identity plan, a splice plan with the
  complete source and target fingerprints, and a reopen plus readback of the
  composed target. That owner was the route before record 0357 too. An
  instrumented build split the 520 µs commit into about 500 µs in the native
  splice and 24 µs in this family's package adoption. The adoption is the only
  new cost.
- The 0627 `tools/perf-baseline` `ole2_range_source` PPT cases still compile
  and run unchanged. Run on the same deck, 5 ABBA rounds, 20 warmups and 50
  samples per run, with the 0627 transport (1 ms per request, 100 MiB/s, 64 KiB
  maximum physical read), after/before p50 is 1.001
  (`ppt_range_source_open`, 18.27 ms), 1.001 (`…_open_one_shape_text`,
  57.69 ms), 0.999 (`ppt_owned_source_control_open`, 179.6 µs) and 1.000
  (`…_open_one_shape_text`, 304.5 µs). The read counts and bytes are
  identical: 14 / 51 / 10 / 47 reads, and 344,064 / 360,875 bytes.

### 7.3 Master order and custom table styles on the primary edit (`937ef82791`)

These are `package::edit::Edit` verbs, staged on the native slide-order
transaction's `document_structure::Transaction`:

- `move_master(&mut self, from: Position, destination: Position) -> Result<()>`,
  `move_master_id(&mut self, master_id: u32, destination: Position) -> Result<()>`,
  `reorder_masters(&mut self, order: &[Position]) -> Result<()>`;
- `insert_custom_table_styles(&mut self, atom: &[u8], CustomTableStylesPlacement) -> Result<()>`,
  `remove_custom_table_styles(&mut self) -> Result<Vec<u8>>`,
  `replace_custom_table_styles(&mut self, atom: &[u8]) -> Result<Vec<u8>>`,
  `move_custom_table_styles(&mut self, CustomTableStylesPlacement) -> Result<()>`.
  The atom is the complete `RoundTripCustomTableStyles12Atom` record, header
  included;
- `set_slide_non_outline_data(&mut self, position: Position, present: bool) -> Result<()>`.

`package::edit` re-exports `CustomTableStylesPlacement`. Each verb records a
`StructuralChange::Document`. The native commit therefore publishes the
change with the live `DocumentContainer`, and its slide-list and
persisted-slide round-trip checks still run unchanged. The structural rebase
(used when media or live-document changes are staged too) replays the change,
and `inverse` inverts it. Masters keep their `masterId`, and slides, notes
and title masters reference masters by that id, so every reference keeps its
master. Refusals are typed: `ElementNotFound` (master, custom table styles),
`InvalidElementValue` (a non-permutation, or bytes that are not exactly one
such atom), `UnsupportedElement` (a second atom), and `Signed`. These edits
had no durable vocabulary, so `to_durable` refused them with
`UnsupportedPatch`, and they did not `join`. Record
[0357b](0357b-ppt-block-roots-and-durable-structure-edits.md) §5–§6 adds
both. Three-way plans report
`document/master-list`, `document/custom-table-styles` and
`slide-list/position:N/non-outline-data`.

Tests (`tests/document_structure_edits.rs`, 8) use POI `slide_master.ppt`
(two main masters and a title master) and `45543.ppt` (two masters and a
custom table-style atom after the `EndDocumentAtom`). They cover reordering
by position, id and permutation, with replay, inverse, `PatchConflict`,
unchanged slide texts and master inventory. They also cover style replace
(another length), move, remove, insert and remove-plus-reinsert (a net
no-op), `fNonOutlineData`, composition with a slide move, the structural
rebase with a media-path change (`unified-ops/0294/video.ppt`), no-op byte
identity, refusals, and the signed rule.

### 7.4 Gates

Each of the three commits passed, with exit 0: `cargo fmt --all --check`;
clippy `--all-features` and `--no-default-features` with `-D warnings`;
`cargo test -p litchi-ppt --all-features` (1,546, 1,553 and 1,561 tests);
`cargo doc` with `-Dwarnings`; `check_api_glossary.py` `ppt-minimal` and
`ppt-full`; `cargo check -p litchi-ops --all-features --all-targets`; and the
perf-baseline `cargo check --locked --all-targets`. After `937ef82791`, the
litchi-ops PPT tests (`--lib`, `crud_matrix`, `ppt_unified_read`,
`ppt_unified_write`, `ppt_elements`) pass: 988 tests, no re-bless. The
litchi-ops `error::` tests pass: 62.

## Follow-ups

- 0357a-F1: a changed scoped commit hashes its target once (about 21 µs on
  the 54 KiB probe deck, +10.7% over the whole-package patch); computing the
  target digest lazily, on the first `apply` of the inverse, would remove it
  from the commit path. Closed by [0357b](0357b-ppt-block-roots-and-durable-structure-edits.md)
  (incremental block roots; the commit is 0.970× the pre-0357a base).

## Primary-route ABBA after the 0357a changes

The primary edit gained verbs and an error variant (§7), so the record 0357
§7 routes were measured again: trunk `5f6fb75164` against this branch at
`cf2088fa00`, release, `taskset -c 29`, binary-level ABBA (single 61, batch
15, elements 21 rounds), host load 7–10 from other agents. Every after/before
ratio is 0.96–1.03 except the first 10k text batch (native 1.072, unified
1.036); three more batch-only ABBA runs put it at native 1.025 and unified
0.998, i.e. noise. Pair ratios are unchanged (single 1.02–1.04, batch
1.09–1.10, the two microsecond reads 1.42 and 3.1 as in record 0323).
Evidence: [`results/0357a/primary/`](results/0357a/primary/) (`summary.json`,
the per-leg JSON, `rb-*.json` for the repeat batch runs, binary SHA-256).
