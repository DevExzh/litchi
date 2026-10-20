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
