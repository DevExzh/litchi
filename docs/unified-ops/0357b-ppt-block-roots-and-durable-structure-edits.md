# 0357b — incremental scoped-patch digests and durable structure edits

- Status: implemented on `uo2/w3-ppt` (follow-up of record
  [0357a](0357a-ppt-record-sized-patches.md)).
- Date: 2026-10-20.
- Input: the coordinator's two items: close 0357a-F1 (a changed scoped
  commit must be back within 5% of the pre-0357a base `e673d4addb`, with
  exact apply and inverse and the `PatchConflict` check kept), and give the
  primary-edit verbs of 0357a §7.3 a durable form and `join` support.
- Owned: `crates/litchi-ppt/src/scoped.rs`, `crates/litchi-ppt/src/scoped/digest.rs`
  (new), `crates/litchi-ppt/src/package/owner.rs` (the digest cache), the
  probe `crates/litchi-ppt/examples/scoped_patch_probe.rs`,
  `tools/unified_bench_matrix.py` (`p0357b`) and the 0357a annotation in
  `results/0289/bench-matrix-curated.json` (§§1–4);
  `crates/litchi-ppt/src/slide_order.rs` (durable document-structure
  operations, join), `crates/litchi-ppt/src/package/edit.rs` (docs, the
  removed refusal), `tests/durable_patches.rs` and
  `tests/document_structure_edits.rs` (§§5–8).
- `performance_claim`: measured (§3).

## 1. F1: what made the commit slower

Record 0357a bound each scoped patch endpoint by its length and the SHA-256
of the whole image. A changed commit hashed the whole target, 54,784 bytes
on the probe deck, which took about 21 µs of 217 µs. The source digest was
cached on the `Package`, so steady-state commits paid only the target hash.
The record-sized ranges themselves (one linear compare) are cheap.

## 2. The design: block roots, derived incrementally

`scoped::digest` (new) binds an image by its **block root**. The root is
SHA-256 over a domain tag (`litchi-ppt/scoped/block-root/v1`), the image
length (u64, little-endian) and the SHA-256 of every 512-byte block in
order. The block size is the diff's block size, one OLE sector, and the
last block may be shorter. Two images with the same root have the same
length and the same digest in every block. So, short of a SHA-256
collision, they are the same bytes. The root is as exact a binding as one
hash over the whole image. Any package that differs from a patch source or
target in any byte, or in length, is still refused with `PatchConflict`.
`a_root_binds_every_byte_and_the_length` checks first, middle and last
bytes, a one-byte truncation and padding.

Every `Package` owns an `ExactCache` (it replaces 0357a's
`OnceLock<BlobId>`). The cache is shared by clones, and each field is
computed at most once:

- `blocks`: the block-digest list (32 bytes per block). It is computed on
  first use, or seeded by the publisher that produced the package.
- `root`: the block root. It is seeded or computed from `blocks`.
- `recipe`: an ancestor's list plus the block spans in which this package
  may differ from it. A replayed package gets a recipe instead of a list.
  The list is derived from the recipe only when something needs it.
- `signed`: whether the bytes carry a digital-signature component. Before,
  every changed commit parsed the CFB for this twice (in `publish_owner`
  and again in the publisher), and it did so in the base too. It is now
  computed once per package.

The operations work as follows.

- **Commit** (`Base::publish`). Compute the ranges as before. Take the
  source's list from its cache. Derive the target's list by hashing only
  the blocks the ranges cover, the appended blocks, and a tail block whose
  extent changed. Copy every other digest. Compute the target root (one
  hash over 32 bytes per block), then seed the published package with the
  derived list. On the probe deck, a commit hashes 26 of 107 blocks
  (13,312 bytes) plus a 3.4 KB root, instead of 54,784 bytes. The published
  package is opened from the owned target bytes
  (`Package::from_owned_bytes_with_options`, with the same input limit),
  which saves one copy of the image. Debug builds assert that every derived
  list equals the full list.
- **Apply / inverse** (`scoped::apply`). Compare the receiving package's
  root (cached, seeded, or computed once) with the patch source root. Then
  rebuild the bytes as before, and seed the result with the patch's target
  root and a recipe. The recipe is the receiving package's list and the
  patch's spans. If that list is not resolved yet, it is the receiving
  package's own recipe with the union of both span lists. That composition
  is sound: a block outside both span lists is equal, with the same extent,
  in all three images. An apply therefore hashes no block. The inverse
  applied to a replayed package compares two seeded roots and hashes
  nothing.
- **Chains.** A commit on a published or replayed package derives from its
  seeded list, or resolves its recipe by hashing only the recipe's spans.
  Chained commits and applies stay O(changed). A recipe chain that is never
  resolved accumulates spans, so resolving it costs at most one pass over
  the image.
- **First use.** A package that was opened, not published, has no list.
  Its first scoped commit, apply or no-op hashes it once, in 512-byte blocks
  (85 blocks on the probe deck).

Alternative considered: computing the target digest lazily, on the first
inverse or durable use, as 0357a-F1 suggested. The patch does not retain
the target bytes. A lazy digest would therefore have to keep the published
package alive from the patch, or re-derive the digest from a package the
caller might have dropped. The block list gives the same commit saving,
retains 32 bytes per patch endpoint as before, and keeps every patch
self-contained.

Patch representation and public surface are unchanged. The endpoints still
hold a length and a 32-byte digest, and retained bytes per patch are
identical (15,360 on the probe). The digest inside is now a block root,
not a whole-image SHA-256. The family `Patch` docs say "SHA-256 block
root". litchi-ppt gains a direct `sha2` dependency (the workspace version
that litchi-core already uses).

Tests: `scoped::digest::tests` (3) cover root sensitivity, derived lists
against full lists (a mid change, growth past a partial tail, shrinkage,
the empty image), and recipe composition through an unresolved recipe. The
four `scoped::tests` byte-patch tests now use block roots. Every family
test exercises the derived and seeded paths through commit, `apply_x` and
the inverse, and the debug assertions compare every derived list and every
rebuilt target with a full hash.

## 3. Measurements

Probe `scoped_patch_probe` (record 0357a) on `test-data/unified-ops/0294/video.ppt`
(43,520 bytes): 1,000 independent single-operation `external_media`
commits on one source, every patch retained, then each patch applied to
the source and each inverse to that result. The probe now also measures
`first_commit_median_us`: the same edit and commit on 200 freshly opened,
never-edited packages. The open is outside the timed region.

Method: three release builds (no LTO), each built from
`git archive <rev> Cargo.toml crates rust-toolchain.toml .cargo` with the
same `Cargo.lock` and probe source, into a separate target directory:

- `base`: `e673d4addb`, the pre-0357a whole-artifact patch;
- `a`: `7d22d1a18b`, the 0357a tip (HEAD before this record);
- `b`: this record, the staged tree `ff8ebf207d` (its `crates/` subtree is
  the F1 commit's).

The runs are binary-level, in ABC-CBA order × 20 rounds, under
`taskset -c 29`. That is 40 runs per build. Peak RSS is from
`/usr/bin/time -f %M`. Times are pooled medians of each run's
per-operation median. Load average was 10.0–11.3, from other agents'
builds on other CPUs. The binaries are not committed. Their SHA-256 values
are in [`results/0357b/binaries.sha256`](results/0357b/binaries.sha256),
the raw runs are in [`abba.tsv`](results/0357b/abba.tsv), the summary is in
[`summary.json`](results/0357b/summary.json), and the scripts are
[`abba.sh`](results/0357b/abba.sh) and [`summarize.py`](results/0357b/summarize.py).
The build directories and their targets were deleted afterwards.

| Metric | base | 0357a | 0357b | 0357b/base | 0357b/0357a |
|---|---:|---:|---:|---:|---:|
| commit, median µs | 194.6 | 215.8 | 188.8 | **0.970** | 0.875 |
| 1,000 commits, total ms | 198.8 | 223.1 | 193.3 | 0.973 | 0.867 |
| apply on source, median µs | 23.3 | 20.6 | 17.5 | 0.754 | 0.851 |
| inverse apply, median µs | 20.8 | 20.5 | 21.5 | 1.034 | 1.050 |
| first commit, never-hashed package, µs | 192.8 | 233.7 | 213.1 | 1.105 | 0.912 |
| bytes retained per patch | 98,304 | 15,360 | 15,360 | 0.156 | 1.000 |
| peak RSS, KiB (median) | 111,092 | 19,710 | 19,712 | 0.177 | 1.000 |

- The changed commit is 0.970× the pre-0357a base. It was 1.109× at
  0357a, and the coordinator's bound is 1.05×. The saving comes from three
  sources: about 15 µs from hashing 13 KB plus the root instead of 54 KB,
  one CFB signature parse per commit (now cached on the package), and one
  image copy.
- Apply on the source got faster: it hashes nothing and parses the
  signature state once per source. The inverse ranges overlap across all
  three builds (base 18.5–24.8, 0357a 17.8–28.1, 0357b 17.4–30.2 µs). The
  inverse does the same work as at 0357a: two seeded roots compared, no
  hash. It still parses the replayed package's signature state once,
  because that package is new.
- The first commit on a never-hashed package pays the source's block list
  once: 213 µs, about 24 µs over the steady state. The 0357a tip paid two
  whole-image hashes there (234 µs). The bench-matrix row has no budget,
  because the cost is paid once per opened package.

The bench matrix gains `p0357b`, whose rows continue the 0357a series
(0357b against the same pre-0357a base). The 0357a commit row is now
`miss (superseded by 0357b)`, and its curated annotation says F1 is closed.

## 4. Gates (F1 commit)

All exit 0 on the F1 tree:

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ppt -p litchi-ops --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ppt --all-targets --no-default-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --all-targets --no-default-features --features ppt -- -D warnings` | 0 |
| `cargo test -p litchi-ppt --all-features` (1,564 tests) | 0 |
| `cargo test -p litchi-ops --all-features --lib --test crud_matrix --test ppt_unified_read --test ppt_unified_write --test ppt_elements --test ppt_coalesced_set --test ppt_create --test ppt_slide_add --test ppt_unified_structure --test ppt_unified_text --test corpus_fidelity` (1,024 tests) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ppt` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/api-glossary` (44 profiles) | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_check_api_glossary tools.test_unified_bench_matrix` (1,304 tests) | 0 |
| `python3 tools/officecli_capability_delta.py --check` | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `python3 tools/check_feature_matrix.py` (31 builds) | 0 |

## 5. Durable form for the document-structure verbs

Record 0357a §7.3 added these verbs to `package::edit::Edit`:
`move_master`, `move_master_id`, `reorder_masters`, the four
custom-table-style verbs and `set_slide_non_outline_data`. Its
`Patch::to_durable` refused all of them. They now have operations in the
native `litchi-ppt` durable format (`slide_order::Patch::to_durable` /
`Snapshot::apply_durable`). The existing operations and their payloads
are unchanged byte for byte. The new ones are additions to the
vocabulary, and the `slide_order` and `durable_patches` tests that pin the
old payloads pass unchanged.

The edits are structural changes (`StructuralChange::Document`), so they
are encoded and replayed like slide moves, in the same structural segment
and in staging order. Every operation carries the slide-list operations'
two preconditions:

- `artifact_sha256`, the exact structural artifact;
- `order_sha256`, the content-free slide-order fingerprint at the point
  the edit was staged.

The staged change now records that fingerprint. A document edit leaves the
slide order alone, so the same fingerprint binds the forward and the
inverse operation. Where the change carries the state it replaces, one
more precondition binds that state:

| Verb | `op` | `target` | extra precondition | `value` |
|---|---|---|---|---|
| `move_master`, `move_master_id`, `reorder_masters` | `document-structure.master-order.set` | `document` | — | the permutation, as zero-based master positions |
| `insert_custom_table_styles` | `document-structure.custom-table-styles.insert` | `document` | — | `{"placement", "record": <blob id>}` |
| `remove_custom_table_styles` | `document-structure.custom-table-styles.remove` | `document` | `styles_sha256` (the removed atom) | `null` |
| `replace_custom_table_styles` | `document-structure.custom-table-styles.replace` | `document` | `styles_sha256` (the replaced atom) | `<blob id>` of the new atom |
| `move_custom_table_styles` | `document-structure.custom-table-styles.move` | `document` | `placement` (before) | the new placement |
| `set_slide_non_outline_data` | `document-structure.slide-non-outline-data.set` | `slide:N` | `non_outline_data` (before) | the flag |

Placements are `before-end-document` and `after-end-document`. An atom is
the complete `RoundTripCustomTableStyles12Atom` record, which travels as a
blob of the durable bundle, as slide payloads do. A bound with no blob
room is `Error::Durable`. Master positions are zero-based; no `masterId`
or persist identifier is serialized. Replay first checks the artifact and
order preconditions, then stages the same `slide_order::Transaction` verb
(the verbs' own validation and the `Signed` rule apply), and checks the
extra precondition against the state the verb replaced. A mismatch, an
unknown vocabulary or a malformed value is an `InvalidFormat` package
error, as for the other durable operations. The durable inverse is the
inverse change, bound to the target-side structural artifact.

`package::edit::Patch::to_durable` no longer refuses these verbs. A
patch replayed from a durable patch still has no operations of its own and
is still refused with `UnsupportedPatch`.

## 6. Join

`package::edit::Patch::join` (native `slide_order::Patch::join`) took only
slide-move patches. It now joins two same-source patches that commute.
Either patch may hold any of these:

- slide moves;
- master reorders;
- custom table-style edits;
- `fNonOutlineData` changes;
- the formatting edits: shape and notes text, anchors, slide visibility,
  advance and visual transition, and external media path and playback.

The joined patch first stages both patches' formatting changes by value on
the source. A formatting change names a source slide position, and no
structure is staged yet, so the projected position is the source position.
It then stages `self`'s structure, then `other`'s, in one native commit.
The commit counts the operations of both. Its patch is an ordinary commit
patch, so it applies and inverts exactly, converts to a durable patch, and
joins again.

Refusals:

- **`Refusal::JoinConflict { target }`** (new, kind `UnsupportedOperation`):
  both patches write the same semantic target, so their order would decide
  the result. The target is named as `plan_three_way` names it, plus
  `slide:N/notes-text`. The cases are:
  - two master reorders (`document/master-list`): two permutations of
    one list do not commute;
  - two custom table-style edits (`document/custom-table-styles`);
  - the same slide's flag (`slide-list/position:N/non-outline-data`);
  - the same shape's text, anchor or notes;
  - the same slide's visibility, advance or transition;
  - the same media path or playback.
- **`Refusal::UnsupportedPatch`** (kind `UnsupportedOperation`):
  - another source. The package route already returns `PatchConflict`
    for this.
  - positions that meet a moved range. Neither patch's slide-list
    positions (moved ranges and `fNonOutlineData` slides) may meet the
    other's moved positions. Otherwise a position would read a different
    slide in the two patches. Overlapping moves are refused as before.
  - slide insertions and removals. They shift every later position.
  - element record edits (comments, hyperlinks, tables and the other
    record-0323 element owners). They replace persisted records that the
    other patch may rewrite too, and replaying them by value is not
    modeled.
  - an inverse patch that holds both formatting and structure. Its
    formatting targets resolve after its structure.

## 7. Tests

`tests/durable_patches.rs` (17 tests; 6 new, 1 changed), all on the public
route:

- **master reorder** (`slide_master.ppt`, three masters): the operation
  shape (no blob needed). Then `to_durable` → deterministic JSON →
  `from_deterministic_json` → `apply_durable` on the source reproduces the
  commit's bytes. The durable inverse restores the source projection
  (master ids, styles atom, every slide's flag, slide texts), and the
  durable patch is refused on another package.
- **custom table styles** (`45543.ppt`): replace with an atom of another
  length, move, remove, and insert into the deck that has none. Each gets
  the same round-trip, byte and inverse checks. An atom with no blob room
  is `Error::Durable`.
- **`fNonOutlineData`**, and **mixed structure**: a slide move, a master
  reorder, a style move and a flag on the moved slide, in one commit.
  The operation order is checked, then the same round trip.
- **joins**: a master reorder joined with a slide move and with a text set
  (3 operations). The test checks the masters, slide order and text
  projection, that the swapped join order gives the same result, exact
  inverse bytes, and a durable round trip and durable inverse of the
  joined patch. A second join test chains a style move, a master reorder,
  a flag off the moved range and a slide move.
- **conflicting joins**: two master reorders, a style move with a style
  removal, the same flag twice, and the same shape's text twice. Each is
  `JoinConflict` with its target. A flag on a slide the other patch moves
  is `UnsupportedPatch` in both join orders.
- **changed**: `joins_refuse_overlap_other_operations_and_other_sources`
  now refuses a slide insertion as the "other operation", and shows that a
  slide-visibility change off the moved range joins.

`tests/document_structure_edits.rs`:
`document_edits_compose_with_slide_moves_and_plan` now expects
`to_durable` to succeed (three operations) and the two master reorders to
be `JoinConflict { target: "document/master-list" }`.

No performance claim is made for §5–§6. The commit path of the primary
edit gains one slide-order fingerprint per staged document edit, which
only these verbs pay.

## 8. Gates (durable and join commit)

All exit 0 on the second commit's tree: `cargo fmt --all --check`;
clippy `-p litchi-ppt -p litchi-ops --all-targets --all-features`,
`-p litchi-ppt --all-targets --no-default-features` and
`-p litchi-ops --all-targets --no-default-features --features ppt`, each
with `-D warnings`; `cargo test -p litchi-ppt --all-features` (1,570
tests); the litchi-ops test list of §4 (1,024 tests, no re-bless);
`cargo doc` with `-Dwarnings`; the glossary gate over all 44 profiles; the
seven `tools` unittest modules (1,304 tests); `officecli_capability_delta.py
--check`; `officecli_corpus.py --check`; `unified_bench_matrix.py --check`;
and `check_feature_matrix.py` (31 builds).
