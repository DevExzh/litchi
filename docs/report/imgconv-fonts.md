# Audit: `litchi-imgconv` and `litchi-fonts` workspace split

## 1. Scope

Verify behavioral parity between `main` and `refactor/workspace-split` for:

| New                                      | Old (on `main`)   |
| ---------------------------------------- | ----------------- |
| `crates/litchi-imgconv/src/`             | `src/images/`     |
| `crates/litchi/src/images/mod.rs` (shim) | (covered above)   |
| `crates/litchi-fonts/src/`               | `src/fonts/`      |

Specifically investigate the apparent shrink of `images/mod.rs` (355 -> 153 lines)
and the deletion of `src/images/extractor.rs` (731 lines).

## 2. Method

- `git diff main..HEAD --stat` over the four paths to enumerate every changed file.
- For each non-zero-line diff, ran `git diff main..HEAD -- <old> <new>` and read
  every `[+-]` line.
- Verified the deleted `src/images/extractor.rs` (731 LoC) was relocated to
  `crates/litchi-ole/src/extractor.rs` (734 LoC, byte-equivalent body).
- Verified the deleted 355-line `images/mod.rs` is now split between:
  - `crates/litchi-imgconv/src/lib.rs` (189 lines: `convert_blip_to_*`, tests, re-exports)
  - `crates/litchi/src/images/mod.rs` (153 lines: `extract_images_from_*`, `parse_blip_store`, re-export `pub use litchi_imgconv::*`)
- Spot-read: `litchi-imgconv/src/lib.rs`, `litchi/src/images/mod.rs`,
  `litchi-ole/src/extractor.rs` lines 280-321, `litchi-fonts/src/lib.rs`,
  both `Cargo.toml` files, and every file with a non-trivial diff
  (`bse.rs`, `wmf/svg/simd.rs`, `wmf/svg/mod.rs`, `wmf/mod.rs`, `emf/mod.rs`,
  `emf/svg/state.rs`, `emf/svg/converter.rs`, `pict/mod.rs`, `blip.rs`,
  `svg_utils.rs`, fonts `loader.rs`, fonts `subsetter.rs`).

## 3. Findings

### `crates/litchi-imgconv/`

OK. All file diffs fall into the allowed buckets:

- **Path moves**: every EMF/WMF/PICT/SVG record-parser/converter file moved
  with zero non-import changes (verified: `emf/records/{bitmap,drawing,mod,objects,path,state,text,types}.rs`,
  `emf/svg/{mod,path,buffer}.rs`, `pict/{types,parser,data,converter}.rs`,
  `wmf/{constants,parser,converter,svg/{state,style,bounds,renderer,transform}}.rs`,
  `svg.rs` — diffs are all `use crate::common::*` -> `use litchi_core::*` or
  `use crate::images::*` -> `use crate::*`).
- **Doc-link updates**: `litchi::images::wmf::convert_wmf` etc. -> `litchi_imgconv::...`
  inside `///` examples — 100% of the substantive content of `emf/mod.rs`,
  `wmf/mod.rs`, `wmf/svg/mod.rs`, `pict/mod.rs`, `svg_utils.rs` diffs.
- **Inlined helper**: `Blip::try_from_escher_record` (which referenced
  `EscherRecord` from the old `crate::ole`) was deleted from `blip.rs` and the
  byte-reconstruction body was inlined verbatim into the only call site,
  `litchi-ole/src/extractor.rs:295-306`. The reconstruction (8-byte header +
  data) is byte-identical and the outer match arm guards the same seven
  `BlipEmf|BlipWmf|BlipPict|BlipJpeg|BlipPng|BlipDib|BlipTiff` discriminants.
  The error path differs only in wording ("Record type 0x... is not a supported
  image record" vs "Not a BLIP record"), which is allowed.
- **One stylistic regression on a `bse.rs` validation, not a logic change**:
  `crates/litchi-imgconv/src/bse.rs:130` — `!name_len.is_multiple_of(2)` was
  rewritten as `name_len % 2 != 0`. Same boolean, same Err arm, no logic drift.
- **`wmf/svg/simd.rs`**: only adds an explanatory comment plus
  `#[allow(clippy::incompatible_msrv)]` over the AVX-512 path; no codegen
  change.

Representative no-op moves:
`crates/litchi-imgconv/src/emf/records/{bitmap,drawing,objects,path,state,text,types}.rs`
(0 lines changed each).

### `crates/litchi/src/images/mod.rs` (umbrella shim, 153 lines)

OK. Compared line-by-line against the relevant block of the old 355-line
`src/images/mod.rs`:

| Old `src/images/mod.rs` block                 | New location                               |
| --------------------------------------------- | ------------------------------------------ |
| Module declarations, `pub use blip::...`, `pub use bse::...`, all `convert_blip_to_*` helpers, tests | `crates/litchi-imgconv/src/lib.rs`         |
| `extract_images_from_ppt`, `extract_images_from_doc`, `extract_images_from_escher`, `parse_blip_store`, doc-comments | `crates/litchi/src/images/mod.rs` (verbatim except `crate::common::error::*` -> `litchi_core::error::*` and `crate::ole::OleFile` unchanged) |
| `pub use extractor::{ExtractedImage, ImageExtractor}` | `pub use litchi_ole::extractor::{ExtractedImage, ImageExtractor}` |

One feature-gating tightening worth noting (allowed, not flagged):

- On `main`, `extract_images_from_escher` and `parse_blip_store` were
  unconditionally compiled (their bodies called the always-present
  `ImageExtractor`). On the new branch they are gated
  `#[cfg(all(feature = "ole", feature = "imgconv"))]` because
  `ImageExtractor` now lives in `litchi-ole`. With the default feature set
  (`ole + ooxml + ooxml_encryption + eval_engine` plus the umbrella's default
  `imgconv`+`fonts` per `crates/litchi/Cargo.toml`), both functions remain
  reachable. No call-graph regression.

No 200-line "vanish" — every line of the old 355-line file is accounted for in
the two new files, modulo doc-comment de-duplication.

### `src/images/extractor.rs` -> `crates/litchi-ole/src/extractor.rs`

OK. The 731 -> 734 line file is functionally identical:
- Imports rewired (`crate::common::error` -> `litchi_core::error`,
  `crate::images::*` -> `litchi_imgconv::*`, `crate::ole::*` -> `crate::*`).
- Three `#[cfg(feature = "imgconv")]` removals on `to_png`/`to_jpeg`/`to_svg`
  helpers — these now live in a crate where `imgconv` is unconditionally a
  dependency, so the gate is redundant. No body changes.
- One `#[cfg(feature = "ole")]` removal at module level for the OLE-helpers —
  same reasoning (the file is inside `litchi-ole`).
- The Escher-record -> `Blip` reconstruction inlined from the deleted
  `Blip::try_from_escher_record`, with byte-for-byte identical header packing
  (`(instance << 4) | version` LE, `record_type_raw` LE, `length` LE, raw
  payload).

### `crates/litchi-fonts/`

OK.
- `lib.rs` is the renamed `src/fonts/mod.rs` with the per-item
  `#[cfg(feature = "fonts")]` gates removed (the entire crate is the gate now).
  The umbrella `crates/litchi/Cargo.toml:69` still gates the public surface
  with `fonts = ["dep:allsorts", "dep:litchi-fonts"]` and
  `crates/litchi/src/lib.rs:342` still has `#[cfg(feature = "fonts")] pub mod
  fonts { pub use litchi_fonts::*; }`. End-user feature semantics unchanged.
- `loader.rs`: only `use crate::fonts::{...}` -> `use crate::{...}`. Font
  matching, fallback, and `font-kit::SystemSource` selection logic untouched.
- `subsetter.rs`: only `use crate::fonts::{...}` -> `use crate::{...}`.
  Allsorts `subset(CmapTarget, SubsetProfile, ...)` call unchanged.
- New `Cargo.toml` declares the same dependency set (`allsorts`, `font-kit`,
  `roaring`, `thiserror`) that was already in the workspace.

## 4. Summary

**Verdict: PURE MECHANICAL.**

No EMF/WMF/PICT decoding algorithm change, no removed conversion path, no
default-DPI / dimension change, no validation removed, no `panic!`/`assert!`
removed, no new `TODO`/`FIXME`/`unimplemented!()`, no font-matching change.
Every byte of the 355-line `images/mod.rs` and the 731-line `extractor.rs` is
relocated, with diffs limited to import paths, doc-link rewrites, redundant
feature-gate removals, the `try_from_escher_record` inline-and-delete, and one
`is_multiple_of` -> `% 2 != 0` style swap.
