# 0127 — PPTX canonical constructor

## Public boundary

`litchi_pptx::Package::new()` now returns the finalized minimal presentation.
It supports canonical reads, immutable edits, replay and stream publication
without an authoring-flush ZIP round trip. The existing graph constructor is
private `new_authoring`: public `new()` finalizes immediately, while the builder
retains that authoring state until `build()` finalizes and releases the projection.
One private rollback regression also exercises the internal authoring state.
Unified PPTX creation now invokes the canonical native constructor directly.

The public `presentation_mut` entry point is deleted. Automatic font configuration
is available on `Builder::font_embedding` before `build`; the package's setter is
private and its consuming `with_font_embedding` wrapper is deleted. Both PPTX
public-API glossary profiles reject all three removed names. No compatibility
shim or deprecated alias is introduced.

Native authoring fixtures and public doctests now use `Builder`. Reopens that
establish physical source provenance or exercise serialization are retained.
Tests asserting the deleted public dirty-writer/hydration state now exercise
canonical edits, no-op preservation and finalized content instead. The one
internal rollback regression explicitly constructs private authoring state.
The fresh-package re-encryption refusal is now `EncryptionPolicy` (no retained
profile), with an empty sink, rather than the obsolete authoring-state refusal.

A new public regression changes the slide-master graph directly from
`Package::new()`, verifies original bytes remain unchanged, checks replay and
inverse, and reopens the published changed package. Other migrated regressions
verify notes/tag edits immediately after building and exact designer no-ops.
The immutable-owner status in the glossary is corrected to acknowledge remaining
direct package mutators: completing the public constructor split does not prove
the entire single-write-path policy. Authoring verbs and remaining mutators still
need edit/commit migration; not every `UnsafeEdit` vocabulary entry is removed.

## Verification

The full native encryption/automatic-font suite passed 1,515 tests and doctests
across 97 test targets (three existing ignores). This includes 964 unit tests,
the migrated integration fixtures, the direct fresh-package edit/replay/inverse
regression and 20 doctests. Independent read-only review found no missed live
consumer or constructor-boundary regression. The combined
DOCX/XLSX/XLSB/PPTX/encryption facade suite passed 1,081 tests and doctests.
All-target native/facade Clippy and 20 native doctests passed with warnings denied.
Root examples compiled with DOCX/XLSX/PPTX/encryption/automatic fonts, and the
PPTX-only facade check without encryption passed. Both PPTX glossary profiles
passed: 75 banned names in minimal and 66 in full, with zero transaction families
claimed by either profile.

A native integration run exposed two absent LibreOffice SVG corpus
fixtures. Three required corpus files were temporarily linked from the read-only
base checkout's `3rdparty/libreoffice-core/sd/qa/unit/data/pptx`:

| File | Bytes | SHA-256 |
|---|---:|---|
| tdf163852.pptx | 9579 | 2525665a1e430a9ee543faf456424bd7b994cce0a2d76470a5a07453e3e2ab39 |
| tdf164622.pptx | 18696 | 33b86388804554e35cb3b07e81cf5db5cb6bb02e76716c71f233b57325c2d199 |
| tdf169496_hidden_graphic.pptx | 35202 | f717ca4657257b538466c772eca2a4d4886cb5b919c2ad5f044377fcee124fc8 |

These links are temporary test inputs, not committed artifacts. Corpus checks
only cover the files available to this worktree; no complete native-application
interoperability corpus or visual rendering claim is made. The handout and font
subset adapter gaps recorded in 0126 remain unchanged. This wave does not claim
paired performance acceptance, all-format vocabulary conformance or goal completion.

Verification commands (offline, serialized):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption,automatic-fonts
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption,litchi-pptx/automatic-fonts --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo check -p litchi --no-default-features --features docx,xlsx,pptx,encryption,automatic-fonts --examples
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features pptx
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption,automatic-fonts --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Commit and cleanup

The rewritten base remains an ancestor of HEAD, so no rebase replay is necessary.
Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
Git metadata filesystem. No commit was created. The required staged size/binary
guards printed nothing, and the changed/untracked artifact audit found no forbidden
files. After all verification processes terminated, this checkpoint's temporary
fixture links, `/tmp/litchi-unified-0127-*` scratch, generated root `Cargo.lock`,
`target` and Python cache were removed. Tracked tool lockfiles, pre-existing
fixture directories and inherited worktree changes were preserved. No temporary
worktree was created; delegated review was read-only. The full goal remains active.
