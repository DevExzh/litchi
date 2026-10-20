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
  `results/0289/bench-matrix-curated.json`.
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
