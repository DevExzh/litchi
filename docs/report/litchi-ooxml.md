# litchi-ooxml — Workspace-Split Behavioral-Parity Audit

## 1. Scope

Compare `crates/litchi-ooxml/` (current `refactor/workspace-split`) against
`src/ooxml/` (on `main`), excluding `src/ooxml/opc/` (now `litchi-opc`).

Subsystems audited:

- `docx/` (Word OOXML reader + writer)
- `xlsx/` (Excel OOXML reader + writer; shared-strings, sheet relations)
- `xlsb/` (binary Excel OOXML)
- `pptx/` (PowerPoint OOXML)
- `crypto/` (Standard-2007 + Agile encryption, OLE wrapper)
- Top-level at the audited carve-out: `api.rs`, `error.rs`, `metadata.rs`,
  `custom_properties.rs`, `pivot.rs`, `common/`, `charts/`, `drawings/`,
  `fonts/`. Shared custom properties have since moved to
  `litchi-ooxml-common::custom`.

Carve-out commit: `41834df` (P4d). Companion commits `3c1d141` (doctest
paths), `71158bd` (READMEs).

## 2. Method

- File-set diff: `git ls-tree -r main src/ooxml/` vs
  `find crates/litchi-ooxml/src -type f`. Only delta is
  `mod.rs` → `lib.rs` (expected) and the `opc/*` files (carved out, not
  leaked).
- Per-subsystem `git diff main..HEAD --stat` to bucket churn.
- Spot-check 3 files per subsystem; one writer + one reader per format.
- Heuristic counters: TODO/FIXME (15 → 15), production
  `panic!`/`assert!`/`debug_assert!` (5 → 5).
- OPC-leak grep: no `opc/` directory in `litchi-ooxml`; only the shim
  `pub mod opc { pub use litchi_opc::*; }` in `lib.rs:48`.

## 3. Findings

### docx — ✅ pure mechanical

Representative files:
`crates/litchi-ooxml/src/docx/parts/document_part.rs`,
`crates/litchi-ooxml/src/docx/header_footer.rs`,
`crates/litchi-ooxml/src/docx/writer/doc.rs`, `docx/mod.rs`,
`docx/package.rs`.

Changes: `crate::ooxml::` → `crate::`; `crate::common::*` →
`litchi_core::*`; `crate::ooxml::opc::*` → `litchi_opc::*`;
`crate::fonts::*` → `litchi_fonts::*`; doc-link path updates
(`litchi::ooxml::docx::Package` → `litchi_ooxml::docx::Package`).

Notable seam:
`crates/litchi-ooxml/src/docx/mod.rs:115-126` adds a crate-local
`DocxElement` enum to avoid `crate::document::DocumentElement`
(reverse dep on umbrella). Used in `document_part.rs:268-360` and
`document.rs:220`. Body of the elements parser is byte-identical to
`main:src/ooxml/docx/parts/document_part.rs` aside from the variant
name; the umbrella translates `DocxElement::{Paragraph,Table}` into
`DocumentElement::Paragraph(Paragraph::Docx(_))` /
`DocumentElement::Table(Table::Docx(_))` at the seam (per commit
message; verified via grep — no logic moved into the wrapper).

Stylistic clippy refactors (semantically identical): nested
`if e.local_name().as_ref() == b"…" { … }` inside match arm rewritten
as a `match-guard`, e.g. `header_footer.rs:118-129` and several
sibling files. Both forms fall through to `_ => {}` on mismatch and
do nothing — equivalent.

### xlsx — ✅ pure mechanical (shared-strings, sheet-rels OK)

Spot-checks:

- `crates/litchi-ooxml/src/xlsx/shared_strings.rs` — only import
  change `crate::sheet::Result` → `litchi_core::sheet::Result`.
  All capacity constants and parsing logic byte-identical.
- `crates/litchi-ooxml/src/xlsx/workbook.rs` — sheet/relationship
  resolution (`load_print_settings`) unchanged; encryption gate
  rename `feature = "ooxml_encryption"` → `feature = "encryption"`
  (per Cargo.toml feature carve-out).
- `crates/litchi-ooxml/src/xlsx/writer/sheet.rs` — drawing/image
  embedding (`write_a_blip_embed_rid_num`, `write_a_stretch_fill_rect`)
  paths unchanged; only crate prefix shifts.
- `crates/litchi-ooxml/src/xlsx/parsers/workbook_parser.rs`,
  `worksheet_parser.rs`, `pivot/reader.rs`, `pivot/writer.rs`,
  `styles/parser.rs` — pure import rewrites.

One micro-refactor at `xlsx/workbook.rs:368-378`: nested
`if rows.is_none()` inside match arm → match-guard
`Some(ch) if ch.is_ascii_digit() && rows.is_none() =>`. When the
guard fails, fall-through hits `_ => {}` (no-op), identical to the
old empty-`if` body. Verified equivalent.

`api.rs` line 187 changed `rust,no_run` → `ignore` for the doc-test
example that referenced `litchi::sheet::WorkbookTrait` (no longer in
crate scope). Doc-test visibility, not runtime.

### xlsb — ✅ pure mechanical

Spot-checks:

- `crates/litchi-ooxml/src/xlsb/cells_reader.rs` — same nested-if →
  match-guard refactor for record-type dispatch (records `0x0001`
  through `0x000B`). Each guard tests the same `buf.len() >= N`
  precondition; on failure the arm doesn't match and the `_ => {}`
  arm runs (no-op), identical to the original empty-`if` body.
  Record bodies (BrtCellBlank, BrtCellRk, BrtCellError,
  BrtCellBool, BrtCellReal, BrtCellSt, BrtCellIsst, BrtFmlaString,
  BrtFmlaNum, BrtFmlaBool, BrtFmlaError) are byte-identical
  payload extraction.
- `crates/litchi-ooxml/src/xlsb/writer/worksheet.rs` — only
  import path changes.
- `crates/litchi-ooxml/src/xlsb/error.rs` — `crate::ole::*` →
  `litchi_ole::*` per commit message; logic unchanged.
- Test value change at `xlsb/writer/worksheet.rs:1432-1433`:
  `3.14f64` → `1.5f64` to silence `clippy::approx_constant`
  (per commit message). Test-only; production code unaffected.

### pptx — ✅ pure mechanical

Spot-checks:

- `crates/litchi-ooxml/src/pptx/presentation.rs` — only import
  rewrites and doc-link updates (e.g. `litchi::ooxml::pptx::Package`
  → `litchi_ooxml::pptx::Package`); slide and slide-master
  resolution via `PresentationPart` and `PackURI::from_rel_ref`
  unchanged.
- `crates/litchi-ooxml/src/pptx/parts/slide.rs` — same
  nested-if → match-guard refactor in `name()` and `text()`
  (cSld/`a:t` element handling). Equivalent.
- `crates/litchi-ooxml/src/pptx/writer/slide.rs`,
  `pptx/writer/pres.rs`, `pptx/shapes/{base,picture,table,
  textframe}.rs` — pure import rewrites.

### encryption (crypto) — ✅ pure mechanical

`crates/litchi-ooxml/src/crypto/agile.rs`,
`standard2007.rs`, `ole_encrypted_package.rs`, `mod.rs`.

Changes:

- `crate::ole::*` → `litchi_ole::*`
  (`ole_encrypted_package.rs:1-3`, `crypto/mod.rs:8`); the
  `crate::ole::writer::OleWriter` import becomes
  `litchi_cfb::writer::OleWriter` per `Cargo.toml` (litchi-ole
  re-exports it; both crates resolve to the same type).
- MSRV-clean rewrite (per commit message):
  `!encrypted_hmac_key.len().is_multiple_of(AGILE_BLOCK_SIZE)`
  → `encrypted_hmac_key.len() % AGILE_BLOCK_SIZE != 0`
  at `agile.rs:694, 726, 760, 799` and
  `standard2007.rs:347`. Numerically identical (both reject
  empty + non-block-aligned ciphertext).
- Constants (`AGILE_BLOCK_SIZE = 16`, AES key sizes, salt-block
  derivation), HMAC verification flow, AES-CBC mode wiring,
  password-hash spin count loop, and key-derivation paths are all
  byte-identical to `main`. `EncryptionMode` enum unchanged.

### top-level — ✅ pure mechanical

- `api.rs` — re-export prefix change `crate::ooxml::*` →
  `crate::*`; one doc-test marker `rust,no_run` → `ignore`
  (mentioned above).
- `error.rs` — adds an `impl From<OoxmlError> for litchi_core::Error`
  (lines 63-83) to satisfy the orphan rule. Mapping is exhaustive
  (every `OoxmlError` variant translated). Mentioned explicitly in
  the carve-out commit. No new error semantics: each case maps to
  the analogous `litchi_core::Error` variant that the umbrella
  previously produced via the umbrella's own `From` impl in
  `src/error_ext.rs` (now removed per commit notes).
- `pivot.rs` — 100% rename, no content changes.
- `common/properties.rs` — single import shift
  `crate::common::xml::escape_xml` → `litchi_core::xml::escape_xml`.
- `metadata.rs`, historical `custom_properties.rs` —
  `crate::common::Metadata` → `litchi_core::Metadata`; OPC paths via
  `litchi_opc`. The latter is no longer owned by this migration host.
- `drawings/{blip,ext,fill,xfrm,mod}.rs` — only `blip.rs` has an
  import change; others byte-identical.
- `charts/` — imports only, plus one stylistic refactor at
  `charts/reader.rs:163-181` (`parse_wall_floor`): nested-if
  inside Start/Empty match arm rewritten as match guard
  `if e.local_name().as_ref() == b"c:thickness" =>`. Same fall-
  through semantics.
- `fonts/{mod,obfuscation}.rs` — feature-gated paths
  `crate::fonts::*` → `litchi_fonts::*`; `crate::common::id::*`,
  `crate::common::encoding::*`, `crate::common::simd::xor::*` →
  `litchi_core::*`.
- No OPC source files leaked (verified by `find` and
  `grep "pub mod opc"`); the only `opc` reference inside
  `litchi-ooxml` is the re-export shim `lib.rs:48-50`.

### Cargo.toml — ✅

`crates/litchi-ooxml/Cargo.toml`:

- Feature `encryption = [aes, cbc, hmac, sha1, litchi-cfb,
  litchi-ole]` matches the umbrella's
  `ooxml_encryption = ["…", "litchi-ooxml/encryption"]`
  in `crates/litchi/Cargo.toml:64`.
- Feature `fonts = [litchi-fonts, allsorts]` matches old
  `fonts` gating.
- `default = []` (empty) is correct: the umbrella `litchi`
  crate owns the default-feature surface
  (`default = ["ole", "ooxml", "ooxml_encryption",
  "eval_engine"]` at `crates/litchi/Cargo.toml:46`).
- Workspace-wide deps (`atoi_simd`, `quick-xml`, `memchr`,
  `rand`, `roaring`, `ryu`, `smallvec`, `bytes`, `chrono`,
  `encoding_rs`, `fast-float2`, `sha2`, `thiserror`,
  `xml-minifier`, `zerocopy`, `base64`, `soapberry-zip`)
  match what was being used through the umbrella.

## 4. Summary

**Verdict: PURE MECHANICAL** (0 logic concerns).

All differences fall into the allowed categories:

1. Import-path rewrites (`crate::ooxml::*` → `crate::*`;
   `crate::common::*` → `litchi_core::*`; `crate::ooxml::opc::*`
   → `litchi_opc::*`; `crate::ole::*` → `litchi_ole::*`;
   `crate::fonts::*` → `litchi_fonts::*`).
2. Doc-comment crate-path updates
   (`litchi::ooxml::*` → `litchi_ooxml::*`).
3. MSRV-driven `is_multiple_of(N)` → `% N != 0` (5 sites in
   `crypto/`).
4. Clippy-driven nested-`if`-in-match-arm → match-guard rewrites
   (verified semantically equivalent because the fall-through
   path leads to a `_ => {}` no-op, matching the old empty
   `if` body).
5. One crate-local `DocxElement` enum in `docx/mod.rs` to avoid
   a reverse dependency on the umbrella; the parser body in
   `docx/parts/document_part.rs` is otherwise byte-identical.
6. Orphan-rule `impl From<OoxmlError> for litchi_core::Error`
   relocated into `error.rs`; mapping is exhaustive and
   preserves prior error-semantics.
7. One xlsb writer test literal `3.14` → `1.5`
   (`clippy::approx_constant`); test-only.
8. One xlsx doc-test marker `rust,no_run` → `ignore`
   (`api.rs`); doc-test visibility only.

No algorithm, parser-branch, validation, bounds-check, struct
layout, encryption-mode, key-derivation, AES/HMAC, sheet-relation,
shared-strings, drawing-embed, theme/style propagation, or
formula-rewrite logic was changed. TODO/FIXME counts are stable
(15 → 15) and production-code panic/assert counts are stable
(5 → 5). No `unimplemented!()` or new `todo!()` introduced.
