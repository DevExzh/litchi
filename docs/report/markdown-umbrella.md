# Workspace-split audit: markdown emission + umbrella crate

## 1. Scope

Compare `refactor/workspace-split` to `main` for:

| New (refactor branch)                    | Old (`main`)                                      |
| ---------------------------------------- | ------------------------------------------------- |
| `crates/litchi-markdown/src/`            | `src/markdown/{config,traits,unicode}.rs`         |
| `crates/litchi/src/markdown/`            | `src/markdown/{mod,document,presentation,writer}.rs` |
| `crates/litchi/src/{lib,document,presentation,sheet}.rs` etc. | `src/{lib,document,presentation,sheet}.rs` etc. (non-eval bits) |
| `crates/litchi/src/metadata_ext.rs` (new, 83 LOC) | inherent method on `Metadata` in `src/common/metadata.rs` |
| `crates/litchi/src/images/mod.rs` (new shim, 153 LOC) | `src/images/mod.rs` (full module on main)         |

Goal: confirm logic is unchanged and the move is mechanical.

## 2. Method

- `git diff main..HEAD` over each file pair.
- Re-read `crates/litchi/src/lib.rs` and `crates/litchi/src/metadata_ext.rs` end-to-end.
- For `markdown/writer.rs` (41-line diff) walked every chunk to confirm only path swaps.
- Cross-checked sibling enums (`DocElement` / `DocxElement`) introduced in `litchi-ole` / `litchi-ooxml` to verify the new umbrella `Document::elements()` wrappers reproduce the old shape.
- Confirmed `Metadata::to_yaml_front_matter` body matches byte-for-byte between old inherent method and new trait impl.

## 3. Findings

### 3.1 `litchi-markdown` (new leaf crate) — OK

Files: `config.rs`, `traits.rs`, `unicode.rs`, `lib.rs` (new 24-line crate root).

- `config.rs` (22-line diff): only doc-comment `use litchi::markdown::...` → `use litchi_markdown::...` updates.
- `unicode.rs` (14-line diff): same — doctest-only path swaps. PHF maps for super/subscript and the four `to_*` / `convert_to_*` / `can_convert_to_*` functions are byte-identical.
- `traits.rs` (10-line diff): `use crate::common::Result` → `use litchi_core::Result`; doctests `no_run` → `ignore` (these doctests reference `litchi::Document`, which is not in `litchi-markdown`'s deps, so they cannot compile here — `ignore` is the right call).
- `lib.rs` (new): pure declarations + re-exports of `MarkdownOptions`, `FormulaStyle`, `ScriptStyle`, `StrikethroughStyle`, `TableStyle`, `ToMarkdown`. No new logic.

### 3.2 `litchi` umbrella (`crates/litchi/src/`) — OK

- `document/{mod,element}.rs`: 100% rename, zero content diff.
- `document/{paragraph,run,table,types}.rs`: only path swaps (`crate::common::*` → `litchi_core::*`, `crate::rtf::*` → `litchi_rtf::*`, `crate::odf::*` → `litchi_odf::*`) plus `test-data` → `../../test-data`.
- `document/doc.rs` (80-line diff): mostly path swaps; one notable refactor at lines 460-495 (see 3.5 below) is a forced consequence of breaking the umbrella ↔ `litchi-ole`/`litchi-ooxml` circular dep — the new `DocElement` / `DocxElement` enums in those crates have the same two-variant `Paragraph(Box<P>) | Table(Box<T>)` shape as `super::DocumentElement`, and the umbrella's match arms wrap each variant 1:1. No semantic change.
- `presentation/{mod,prs,slide,types}.rs`: pure path swaps (`crate::common::*` → `litchi_core::*`, `crate::odf::*` → `litchi_odf::*`).
- `sheet/{workbook,workbook_types}.rs`: pure path swaps. One incidental clippy-style change in a test:
  - `crates/litchi/src/sheet/workbook_types.rs:220` — `format.clone()` → `format` (because `WorkbookFormat: Copy`). Behavior identical.
- `sheet/functions.rs`: 100% rename, no diff.
- `sheet/text/**`: pure path swaps (`crate::common::{BomKind, strip_bom, write_bom}` → `litchi_core::*`).
- `sheet/mod.rs` (8-line diff): converts `pub mod traits;` / `pub mod types;` to `pub use litchi_core::sheet::{traits, types};` and wraps `eval` as `pub mod eval { pub use litchi_eval::*; }`. The corresponding `litchi_core::sheet::traits.rs` is byte-identical to the old `src/sheet/traits.rs` (`diff` returned 0). API surface preserved.

### 3.3 `metadata_ext.rs` (new file, 83 LOC) — flagged as MINOR public-API source-compat break

Logic OK / behaviorally identical, **but the public method moves from inherent → trait method**.

- Old (`main:src/common/metadata.rs:110-122`):
  ```rust
  impl Metadata {
      pub fn to_yaml_front_matter(&self) -> Result<String> {
          if !self.has_data() { return Ok(String::new()); }
          let yaml_string = serde_saphyr::to_string(self).map_err(|e| {
              crate::common::Error::Other(format!("Failed to serialize metadata to YAML: {}", e))
          })?;
          Ok(format!("---\n{}---\n\n", yaml_string))
      }
  }
  ```
- New (`crates/litchi/src/metadata_ext.rs:21-41`):
  ```rust
  pub trait MetadataYaml {
      fn to_yaml_front_matter(&self) -> Result<String>;
  }
  impl MetadataYaml for Metadata {
      fn to_yaml_front_matter(&self) -> Result<String> {
          if !self.has_data() { return Ok(String::new()); }
          let yaml_string = serde_saphyr::to_string(self)
              .map_err(|e| Error::Other(format!("Failed to serialize metadata to YAML: {}", e)))?;
          Ok(format!("---\n{}---\n\n", yaml_string))
      }
  }
  ```
- Body is byte-equivalent: same `has_data()` short-circuit, same `serde_saphyr::to_string`, same `Error::Other(format!("Failed to serialize ..."))`, same `format!("---\n{}---\n\n", yaml_string)` framing.
- **Concern (allowed-by-rationale, but worth noting):** external callers of `metadata.to_yaml_front_matter()` must now `use litchi::MetadataYaml` to bring the method into scope. This is a (small) source-compat break, not a runtime behaviour change, and the rationale (`serde_saphyr` not a dep of `litchi-core`) is documented in the file header. The umbrella's only internal caller (`crates/litchi/src/markdown/writer.rs:1`) is updated to `use crate::MetadataYaml`. No `unimplemented!()` / `todo!()` / removed assertion.

### 3.4 `lib.rs` public API (99-line diff)

Mostly mechanical with **one concern (additive)** and **one concern (cfg-tightening)**.

- `pub mod common;` (loaded from `src/common/`) → inline re-export module:
  ```rust
  pub mod common {
      pub use litchi_core::*;
      #[cfg(any(feature = "ole", feature = "ooxml", feature = "iwa", feature = "odf", feature = "rtf"))]
      pub use crate::detection_smart::{detect_file_format, detect_file_format_from_bytes};
      pub mod detection { /* re-exports DetectedFormat, detect_format_smart, ... */ }
  }
  ```
  Old API surface preserved as long as a parsing feature is enabled (the default config does).

- `pub mod {ole,ooxml,formula,iwa,odf,rtf,fonts}` switched from owning the code to `pub use litchi_*::*`. The `images` module became a shim re-exporting `litchi_imgconv::*` plus a thin `extractor` glue layer (`crates/litchi/src/images/mod.rs:46,52`).

- **Additive**: new line `pub use sheet::Workbook;` (`crates/litchi/src/lib.rs:363`). Was not a top-level re-export on main but was reachable via `litchi::sheet::Workbook`. Pure addition — should not break callers.

- **Concern (cfg-tightening)**: `detect_file_format` / `detect_file_format_from_bytes` re-exports at crate root were unconditional on main:
  ```rust
  // main:src/lib.rs (last block)
  pub use common::{
      FileFormat, Length, PlaceholderType, RGBColor, ShapeType,
      detect_file_format, detect_file_format_from_bytes,
  };
  ```
  Refactor branch:
  ```rust
  // crates/litchi/src/lib.rs:366-375
  pub use common::{FileFormat, Length, PlaceholderType, RGBColor, ShapeType};
  #[cfg(any(feature = "ole", feature = "ooxml", feature = "iwa", feature = "odf", feature = "rtf"))]
  pub use common::{detect_file_format, detect_file_format_from_bytes};
  ```
  With default features these symbols remain at `litchi::detect_file_format(_from_bytes)`. With `--no-default-features` and zero parsing crates they vanish. Acceptable because the underlying `detection_smart` module is itself cfg-gated identically (it cannot exist without at least one parser), but it is a measurable narrowing of the public API in feature-stripped builds.

- Doc-comment additions explaining the new `common` shape and detection split. No removed `pub` items, no signature changes on `Document::open()` / `Presentation::open()` / `Workbook::open()`.

### 3.5 `markdown/writer.rs` (41-line diff) — OK

Purely import path updates. Walked every changed chunk:

- Line 1-9 imports: `super::config::*` → `litchi_markdown::*`; `crate::common::{Error, Metadata, Result}` → `litchi_core::*`; added `use crate::MetadataYaml;` (needed because `to_yaml_front_matter` is now a trait method, see 3.3); and explicit `use crate::document::{Cell, Paragraph, Run, Table};` (was previously the same line in a different order).
- Line 445: `use crate::formula::omml_to_latex;` → `use litchi_formula::omml_to_latex;`.
- Line 654: `use crate::common::VerticalPosition;` → `use litchi_core::VerticalPosition;`.
- Lines 676-714: every `super::config::ScriptStyle::*` → `litchi_markdown::ScriptStyle::*`; every `super::unicode::*` → `litchi_markdown::unicode::*`. Algorithm unchanged: same `Html` / `Unicode` switch, same `<sup>`/`<sub>` fallback path, same `can_convert_to_super/subscript` short-circuit, same `convert_to_super/subscript` call.
- Line 750: `super::config::StrikethroughStyle::Html` → `litchi_markdown::StrikethroughStyle::Html`. The strikethrough handling logic ("HTML strikethrough: must be self-contained per run") is byte-identical.
- Lines 1510-1532: `crate::formula::*` → `litchi_formula::*` for `MathNode`, `latex::LatexConverter`, `omml_to_latex`.
- Lines 1554-1560: `super::config::FormulaStyle::{LaTeX, Dollar}` → `litchi_markdown::FormulaStyle::{LaTeX, Dollar}`. Same `format!("\\({}\\)", ...)` / `format!("${}$", ...)` / `format!("\\[{}\\]", ...)` / `format!("$${}$$", ...)` outputs.

No table style change, no escape-rule change, no list-handling change, no metadata-block change, no new `TODO`/`FIXME`/`unimplemented!()`, no removed bounds checks.

## 4. Summary

**MOSTLY MECHANICAL (2 concerns)**

| # | Concern | Severity | Location |
|---|---------|----------|----------|
| 1 | `Metadata::to_yaml_front_matter` migrated from inherent method to `MetadataYaml` trait method. Byte-for-byte identical output, but external callers must now `use litchi::MetadataYaml`. | Source-compat break; runtime behaviour unchanged | `crates/litchi/src/metadata_ext.rs:21-41` vs `main:src/common/metadata.rs:110-122` |
| 2 | Top-level re-exports `litchi::detect_file_format` / `litchi::detect_file_format_from_bytes` are now cfg-gated on `any(ole, ooxml, iwa, odf, rtf)`. Default-feature builds unaffected; `--no-default-features` builds lose these symbols at crate root. | Public-API narrowing in stripped builds | `crates/litchi/src/lib.rs:368-375` vs `main:src/lib.rs` (final block) |

Additive change worth noting (not a concern, since allowed): `pub use sheet::Workbook;` added at crate root (`crates/litchi/src/lib.rs:363`).

All markdown emission logic (writer.rs / document.rs / presentation.rs / config.rs / traits.rs / unicode.rs) is path-rename only — no algorithm, escaping, list, table, formula, super/subscript, or strikethrough behaviour changed. `Document::open` / `Presentation::open` / `Workbook::open` dispatch tables are preserved 1:1; the only restructure (`Document::elements()` for `.doc`/`.docx`) is a forced consequence of breaking the circular dep and wraps the new `DocElement` / `DocxElement` variants 1:1 onto the existing `super::DocumentElement` shape.
