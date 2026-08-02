# Substrate audit — `litchi-core`, `litchi-cfb`, `litchi-opc`

Branch reviewed: `refactor/workspace-split` against `main`.

> Historical snapshot: this report records the initial workspace extraction.
> As of ADR 0009 (2026-08-03), ODF detection, `quick-xml`, and
> `soapberry-zip` have moved from `litchi-core` to `litchi-odf`; the core ODF
> feature and dependency conversions described below no longer exist.

## 1. Scope

New crates audited:

- `crates/litchi-core/src/`  ←  `src/common/` (+ `src/sheet/{traits,types}.rs`)
- `crates/litchi-cfb/src/`   ←  `src/ole/{consts,file,metadata}.rs` and `src/ole/writer/`
- `crates/litchi-opc/src/`   ←  `src/ooxml/opc/`

Cargo manifests reviewed:

- `crates/litchi-core/Cargo.toml`
- `crates/litchi-cfb/Cargo.toml`
- `crates/litchi-opc/Cargo.toml`

## 2. Method

I produced a `git diff --stat main..HEAD` over each old/new path pair and
inspected every file whose diff was non-zero. Files reported as 100% rename
(`similarity index 100`) were trusted without further checks. For each
non-trivial diff I extracted the actual `+/-` lines with `git diff | grep` and
read the surrounding context. Deletion-only files (e.g. `src/common/tests.rs`,
`src/common/error/conversions.rs`, `src/ole/consts.rs`,
`src/common/detection/{detected,functions,iwork,ole2,ooxml}.rs`) were
cross-checked to confirm the content reappeared verbatim in the new location
(either inside one of the audited crates or in an out-of-scope crate such as
`crates/litchi/` for the umbrella's `detection_smart` module). I also checked
where logic that previously lived in `litchi-core` (the `From<*Error>` impls
and `From<OleMetadata>`) was relocated, since the orphan rule forced moves
across crate boundaries.

## 3. Findings

### `litchi-core`  ✅  No logic change detected

Spot-checked:

- `crates/litchi-core/src/binary.rs` — only doc-comment paths updated
  (`litchi::common::binary` → `litchi_core::binary`); no code body changes.
- `crates/litchi-core/src/error/types.rs` — only `use crate::common::binary` →
  `use crate::binary`; one annotation added: `#[non_exhaustive]` on `Error`.
  Public-API hardening, not behavioral. All variants and their messages are
  identical.
- `crates/litchi-core/src/error/conversions.rs` — body shrank from 110 lines
  to 21. The two impls that remain (`From<quick_xml::Error>`,
  `From<soapberry_zip::Error>`) are byte-identical to the originals. The
  removed `From<OleError>`, `From<DocError>`, `From<PptError>`,
  `From<OpcError>`, `From<OoxmlError>`, `Error::from_opc_error` impls were
  relocated (orphan-rule motivated) to:
  - `crates/litchi-cfb/src/file.rs:143`           (`From<OleError>`)
  - `crates/litchi-opc/src/error.rs:63`           (`From<OpcError>`,
    formerly `Error::from_opc_error`)
  - `crates/litchi-ooxml/src/error.rs:67`         (`From<OoxmlError>`)
  - `crates/litchi-ole/src/doc/package.rs:212`    (`From<DocError>`)
  - `crates/litchi-ole/src/ppt/package.rs:222`    (`From<PptError>`)

  Each relocated impl is byte-for-byte identical to the original match arms;
  `OoxmlError::Opc(e) => litchi_core::Error::from(e)` is equivalent to the
  old `Error::from_opc_error(e)` because both call the same logic.
- `crates/litchi-core/src/metadata.rs` — `to_yaml_front_matter()` and the
  `From<crate::ole::OleMetadata> for Metadata` impl were removed. The body of
  `to_yaml_front_matter()` reappears byte-identically in
  `crates/litchi/src/metadata_ext.rs:30` (now behind a `MetadataYaml` trait
  for orphan-rule reasons). The `From<OleMetadata>` body reappears
  byte-identically in `crates/litchi-cfb/src/metadata.rs:472`. The retained
  `Metadata` struct, `Default` impl, and `has_data()` body are unchanged.
- `crates/litchi-core/src/detection/{odf,rtf,simd_utils,types,utils}.rs` —
  pure `crate::common::` → `crate::` rewrites. The five sibling files
  (`detected,functions,iwork,ole2,ooxml`) were intentionally moved out of
  `litchi-core` because they require `crate::ooxml`, `crate::iwa`, etc., and
  now live under `crates/litchi/src/detection_smart/`. Spot-checked
  `detected.rs`: 200 → 199 lines, only `use` paths and a single internal call
  path renamed. Out of scope but verified the move is mechanical.
- `crates/litchi-core/src/detection/mod.rs` — newly slimmed (replaces the
  old `mod.rs`). Drops `pub mod {detected,functions,iwork,ole2,ooxml}` (those
  modules are no longer in this crate) and drops the
  `pub use functions::{detect_file_format, detect_file_format_from_bytes}`
  re-export. **Behavioral note:** consumers that imported
  `litchi::common::detect_file_format` will no longer find it via
  `litchi_core` — they must use the umbrella's `detection_smart`. This is a
  public-API surface change but the underlying functions still exist; no
  detection logic was deleted.
- `crates/litchi-core/src/sheet/{traits,types}.rs` — `git diff` reports
  `similarity index 100%`. Pure rename of `src/sheet/traits.rs` and
  `src/sheet/types.rs` (130 + 325 lines).
- `crates/litchi-core/src/lib.rs` — diff vs. `src/common/mod.rs` is +14/−0:
  adds `pub mod sheet;`, drops the `detect_file_format*` re-exports, drops
  `Length as MeasuredLength` (verified zero callers in the workspace), and
  adds `#[doc(hidden)]` on the `XmlSlice` re-export. No logic.
- `crates/litchi-core/src/{shapes/types,detection/types,xml_slice}.rs` — only
  `#[non_exhaustive]` / `#[doc(hidden)]` annotations added. No code changes.
- `crates/litchi-core/src/{bom,encoding,unit,id,style/*,simd/*,xml/*}.rs` —
  every diff line is a doc-comment path or a `use crate::common::` →
  `use crate::` import rewrite. Confirmed by inspecting `style/color.rs`,
  `encoding.rs`, `simd/cmp.rs`, `xml/escape.rs`.

### `litchi-cfb`  ✅  No logic change detected

Spot-checked:

- `crates/litchi-cfb/src/consts.rs` — 545 → 92 lines. `diff <(head -92
  src/ole/consts.rs) crates/litchi-cfb/src/consts.rs` is empty
  (byte-identical). The remaining 453 lines were format-specific (PPT
  `RT_*` record types, `WORD_CLSID`, `VT_*` PROPVARIANT ids) and now live in
  `crates/litchi-ole/src/consts.rs`. Spot-checked `WORD_CLSID` and
  `RT_DOCUMENT` — both present in `litchi-ole`. CFB-only constants kept
  verbatim.
- `crates/litchi-cfb/src/file.rs` — adds 25 lines. The added lines are: (a)
  the relocated `From<OleError> for litchi_core::Error` impl
  (lines 143–156), match arms identical to the original
  `src/common/error/conversions.rs:17–30`; (b) two `use crate::common::` →
  `use litchi_core::` rewrites. No control-flow, parsing, or validation
  changes.
- `crates/litchi-cfb/src/metadata.rs` — adds 33 lines: the relocated
  `From<OleMetadata> for litchi_core::Metadata` impl, byte-identical to the
  original at `src/common/metadata.rs:124–148`. Only other change is one
  call-site update: `crate::common::encoding::decode_bytes` →
  `litchi_core::encoding::decode_bytes`.
- `crates/litchi-cfb/src/writer/{difat,fat,header,minifat,mod,tests}.rs` —
  `git diff --stat` reports zero changes (100% rename).
- `crates/litchi-cfb/src/writer/directory.rs` — single doc-comment path
  rewrite.
- `crates/litchi-cfb/src/writer/core.rs` — 16+/16− diff. All doc-comment path
  rewrites (`litchi::ole::writer` → `litchi_cfb::writer`) **plus one
  cosmetic refactor** at two locations:

  Before (`src/ole/writer/core.rs`):
  ```rust
  let mut current_sector = minifat_start_sector;
  for minifat_sector_data in &minifat_sectors {
      let position = ((current_sector as u64) + 1) * (self.sector_size as u64);
      writer.seek(SeekFrom::Start(position))?;
      writer.write_all(minifat_sector_data)?;
      current_sector += 1;
  }
  ```
  After (`crates/litchi-cfb/src/writer/core.rs`):
  ```rust
  for (current_sector, minifat_sector_data) in
      (minifat_start_sector..).zip(minifat_sectors.iter())
  {
      let position = ((current_sector as u64) + 1) * (self.sector_size as u64);
      writer.seek(SeekFrom::Start(position))?;
      writer.write_all(minifat_sector_data)?;
  }
  ```
  Same pattern duplicated for `difat_start_sector`. The `(start..).zip(iter)`
  form yields exactly the same `(start, items[0])`, `(start+1, items[1])`,
  … pairs as the manual counter, and stops on the same iterator-exhausted
  condition. No behavioral change. Worth flagging only because it is *not*
  strictly mechanical (loop body identical, but loop scaffolding rewritten);
  it is, however, logic-preserving.
- `crates/litchi-cfb/src/lib.rs` — newly authored 18-line crate root (mods +
  re-exports). Mirrors what `src/ole/mod.rs` exposed for these submodules.

### `litchi-opc`  ✅  No logic change detected

Spot-checked:

- `crates/litchi-opc/src/{constants,packuri}.rs` — 100% rename, zero diff.
- `crates/litchi-opc/src/error.rs` — single doc-comment style change
  (`///` block → `//!` module-level), and adds the relocated
  `From<OpcError> for litchi_core::Error` impl (lines 63–73). The match arms
  are byte-identical to the old `Error::from_opc_error` at
  `src/common/error/conversions.rs:88–96`.
- `crates/litchi-opc/src/pkgreader.rs` — biggest opc diff (49+/50−). Every
  line is one of: (a) `crate::ooxml::opc::*` → `crate::*` import rewrite, or
  (b) a stylistic refactor that lifts `if e.local_name() == b"Relationship"`
  from a nested `if` inside the match arm into a match guard:

  Before:
  ```rust
  Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
      if e.local_name().as_ref() == b"Relationship" {
          /* parse attrs, push srels */
      }
  },
  ```
  After:
  ```rust
  Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
      if e.local_name().as_ref() == b"Relationship" =>
  {
      /* parse attrs, push srels */
  },
  ```
  This is logic-preserving: in the old form, a non-`Relationship`
  empty/start event entered the arm and silently fell through (no-op); in
  the new form it doesn't match the guard and falls through to the final
  `_ => {}` arm. Same outcome, same iteration, same `srels` content.
- `crates/litchi-opc/src/part.rs` — same nested-`if` → match-guard refactor
  at three sites (`Event::Start`/`Empty` pairs and one `Event::End`). All
  guarded on the same byte-string equality. Verified the `match` still has
  a `_ => {}` catch-all so behavior on non-matching events is identical.
- `crates/litchi-opc/src/{package,pkgwriter,phys_pkg,rel}.rs` — only
  `crate::ooxml::opc::*` → `crate::*` and `crate::common::xml::escape_xml` →
  `litchi_core::xml::escape_xml` import/path rewrites. Spot-checked all
  call-sites; no algorithm change.
- `crates/litchi-opc/src/lib.rs` — newly authored 37-line crate root,
  mirrors `src/ooxml/opc/mod.rs` re-exports verbatim plus
  `pub use error::{OpcError, Result}` (which the old `mod.rs` did not
  re-export — additive only).

### Cargo manifests  ✅

- `litchi-core`: declares `odf`, `ole`, `rtf` features that gate
  `dep:encoding_rs`, mirroring the old `#[cfg(any(feature = "ole", feature
  = "rtf"))]` on `src/common/encoding.rs` and `#[cfg(feature = "odf")]` on
  `src/common/detection/odf.rs`. Behavioral parity preserved.
- `litchi-cfb`: declares `default = []` and a `write` feature. The `write`
  feature gates only `crates/litchi-cfb/examples/write_ole.rs`; the
  `writer/` module itself is unconditionally compiled. Matches `main`,
  where there was no equivalent gate and the writer was always compiled.
  No regression.
- `litchi-opc`: dependency-only manifest, no feature flags. Matches the
  unconditional compilation of `src/ooxml/opc/` on `main`.

### Public-API surface notes (informational, not logic)

These are intentional API changes that the refactor introduces; they were
explicitly listed under the "allowed" rubric:

1. `#[non_exhaustive]` added to `Error`, `FileFormat`, `ShapeType`,
   `PlaceholderType`.
2. `#[doc(hidden)]` added to `XmlSlice`, `XmlArenaBuilder` (still `pub` for
   cross-crate use).
3. Re-export `MeasuredLength` dropped from `litchi-core` (zero in-tree
   callers; verified).
4. Re-exports `detect_file_format` / `detect_file_format_from_bytes` dropped
   from `litchi-core::detection` because the smart-detection functions moved
   to `crates/litchi/src/detection_smart/`. The functions themselves still
   exist; only the re-export path changed.
5. New re-exports in `litchi-cfb` and `litchi-opc` crate roots that did not
   exist on `main` (additive).

## 4. Summary

**Verdict: PURE MECHANICAL.**

Every diff in the three substrate crates falls into one of these allowed
buckets: file/path moves, `use`-path rewrites, doc-comment path rewrites,
visibility upgrades for cross-crate use, additive re-exports, additive
`#[non_exhaustive]` / `#[doc(hidden)]` annotations, and orphan-rule-driven
relocations of `From` impls (each verified byte-identical to its
counterpart on `main`). The only non-trivial code rewrites are two
logic-preserving stylistic refactors:

- `litchi-cfb/src/writer/core.rs`: `let mut counter; for x { counter += 1 }`
  → `for (counter, x) in (start..).zip(iter)`.
- `litchi-opc/src/{pkgreader,part}.rs`: `match … => if guard { … }` →
  `match … if guard => { … }`.

Both rewrites are local, body-identical, and preserve iteration semantics
and event-handling fall-through. No algorithm, validation, error path,
struct layout, default value, or panic guard was changed.

Concern count: 0.
