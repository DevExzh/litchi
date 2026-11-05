# 0057 — Independently retained PPTX scenes and token classification

## Scope

Advance the L-PPTX-R2 retained-scene prerequisite and L-PPTX-R10 token readers
from record 0010. `shape::owned::{Scene, Handle, Handles}` provides an owned
projection over the existing shape reader, with eager `Slide::scene`,
`SlideLayout::scene` and `SlideMaster::scene` entry points. This is a foundation
for R4c, not the unified PPTX adapter or completion of R2's source memo.

The full goal remains active. Source-backed scene access, weighted retention,
related parts, source/execution fences, R6/R7 text and property projection,
PPTX unified dispatch, all-format writes, Python and paired performance
acceptance remain required. Existing borrowed shape APIs and owner vocabulary
still need glossary migration. No iWork source is changed.

## Ownership and behavior

The owned scene retains one `Arc` over the processed XML, shape records and
string arena. An unchanged owner shares its existing `Arc<Vec<u8>>`; a rewritten
MCE owner retains only the processed buffer. Handles are a scene share plus a
checked pre-order position. They hold no package, presentation or slide owner.
The existing bounded reader and part root validation remain authoritative:
there is no second parser or second shape model. A scene is an immutable
snapshot, not a live-source handle.

`owned::Selector` has only `Name(Cow<str>)` and `Position(Position)`. Exact,
case-sensitive names preserve the existing PowerPoint shape-name policy.
Missing names and positions return `Ok(None)`, including `usize::MAX`.
Ambiguity returns `owned::Error::Ambiguous { matches }`, without the authored
name. Native IDs are not new selectors or owned-handle properties. Raw owner
XML is exposed explicitly by `owned::raw::xml`.

`Handle::view` borrows the existing typed leaf `Shape`; that legacy view still
has physical ID/span/XML accessors and content-bearing Debug. This record does
not claim their migration, and the future unified adapter must not publish
those values in ordinary node properties. The new Scene, Handle, Handles and
Selector Debug implementations are content-free.

Iteration owns the scene, is fused, preserves depth-first pre-order, and offers
root and direct-child traversal. Parent handles use checked reader records.
The scene remains one word and the handle two words; tests assert `Send + Sync`
and independently retained lifetime. No performance overhead budget is claimed.

## Token readers and provenance

`SlideLayoutKind` and `PlaceholderKind` now expose schema-order `ALL` arrays
and `from_token`, inverse to `as_str`. They trim only XML whitespace, preserve
case sensitivity and refuse unknown tokens. Missing-attribute defaults remain
the calling reader's responsibility.

Candidate `b9587eebbb` informed the retained-scene design. Only its
`master_layout/model.rs` token-reader diff was applied; the owned projection
and its tests were adapted for ADR0034's names, selection semantics and
redaction policy. Candidate commits/history, measurement evidence and binaries
were not merged. The local base remains
`32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor of HEAD.

## Verification

- Focused owned-scene unit tests: 6 passed.
- Token-filtered leaf tests: 14 passed.
- Eager slide/layout/master parity and post-owner-drop integration test: passed
  on two real PPTX fixtures, including grouped and duplicate-name shapes.
- All-target/all-feature PPTX Clippy with warnings denied: passed.
- Glossary checker Python tests: 21 passed.

The first full leaf run reached native SVG lifecycle tests that require three
untracked fixtures absent from this checkout: `tdf169496_hidden_graphic.pptx`,
`tdf163852.pptx`, and `tdf164622.pptx`. They were temporarily copied from the
base checkout's `3rdparty/libreoffice-core/sd/qa/unit/data/pptx` for verification.

- `CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --all-features`: 872 unit,
  456 integration and 10 doctests passed (1,338 total).
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-pptx --all-features
  --no-deps`: passed.

- `cargo check --workspace --all-features --lib`: passed.
- Minimal and full PPTX glossary profiles: 10 bans each passed.
- The extended layout/master post-package-drop integration check passed; final
  all-feature/all-target Clippy passed after that coverage change.
- Changed Rust files formatted; `git diff --check` passed. The changed/untracked
  artifact audit found no forbidden compiled binaries or oversized restricted
  evidence.
The minimal/full PPTX glossary profiles now prohibit old selector/index/native
ID/XML spellings in the new owned module; they do not claim migration of the
legacy borrowed APIs.

## Commit and cleanup

`git add -A` again failed because
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` cannot be
created on a read-only filesystem. No commit is claimed. Both required staged
artifact checks printed nothing; the old partial index still needs refreshing
before any later commit.
No temporary worktree was created.

The three temporary fixtures and only their newly created directories were
removed. Generated lockfiles and Python cache were removed.
`cargo clean` removed 23,815 files (28.2 GiB); turn scratch was removed
after reducing verification results into this record.
