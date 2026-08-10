# litchi

High-performance Rust library for parsing Microsoft Office, OpenDocument, and Apple iWork file formats with a unified API.

## Overview

`litchi` is the user-facing umbrella crate of the [Litchi workspace](https://github.com/DevExzh/litchi). It auto-detects file formats and delegates parsing to independently owned format crates (`litchi-doc`, `litchi-ppt`, `litchi-xls`, `litchi-docx`, `litchi-pptx`, `litchi-xlsb`, `litchi-xlsx`, `litchi-opc`, `litchi-ooxml-common`, `litchi-odf`, `litchi-pages`, `litchi-keynote`, `litchi-numbers`, `litchi-rtf`, and shared IWA crates). Most users should depend on this crate rather than the format-specific ones. Canonical low-level legacy-format entry points are the standalone `litchi-doc`, `litchi-ppt`, and `litchi-xls` crates; the umbrella exposes `doc`, `ppt`, and `xls` facades only for their enabled features.

Shared OOXML chart and SmartArt grammar is available through the concise
`litchi::drawing::{chart, diagram}` facade when the `drawingml` feature is enabled.
Concrete formats retain their package-specific anchors and relationships.

## Usage

```toml
[dependencies]
litchi = { version = "0", features = ["docx", "pptx", "xlsx"] }
```

```rust
use litchi::{Document, Presentation, Workbook};

fn main() -> Result<(), litchi::Error> {
    let doc = Document::open("report.docx")?;
    println!("{}", doc.text()?);

    let pres = Presentation::open("slides.pptx")?;
    println!("slides: {}", pres.slide_count()?);

    let wb = Workbook::open("data.xlsx")?;
    println!("sheets: {}", wb.worksheet_count());
    Ok(())
}
```

## Entry Points

- `Document::open` — unified Word reader (`.doc`, `.docx`, `.odt`, `.rtf`, `.pages`).
- `Presentation::open` — unified PowerPoint reader (`.ppt`, `.pptx`, `.odp`, `.key`).
- `Workbook::open` — unified spreadsheet reader (`.xls`, `.xlsx`, `.xlsb`, `.ods`, `.numbers`).
- `detect_file_format` / `detect_file_format_from_bytes` — format sniffing without parsing.

## Bounded OOXML Ingestion

`docx`, `pptx`, `xlsx`, and `xlsb` re-export `ReadLimits`, the shared checked
OPC package-ingestion policy. Ordinary constructors use bounded defaults.
Construct a profile from those defaults with `ReadLimits::builder()` and supply
it to the matching contextual API: DOCX and PPTX `Package::*_with_limits`,
XLSX `Package::*_with_limits` or `Workbook::*_with_limits`, and XLSB
`Workbook::new_with_limits`.

The policy bounds compressed input, ZIP member counts, names, directory
metadata, compressed and uncompressed member sizes, materialized OPC parts,
`[Content_Types].xml`, and relationship XML, attributes, targets, events,
depth, and graph traversal. It operationalizes ECMA-376 Part 2 §7.3.6/§10 and
MS-OI29500 §2.1.1749-1752 for hostile input; it is a Litchi safety policy,
not a specification maximum.

```rust
use litchi::docx::{Package, ReadLimits};

let limits = ReadLimits::builder()
    .max_input_bytes(32 * 1024 * 1024)?
    .max_archive_members(10_000)?
    .build()?;
let package = Package::open_with_limits("untrusted.docx", limits)?;
# let _ = package;
# Ok::<(), litchi::Error>(())
```

Macros, VBA, ActiveX, controls, OLE objects, and embedded code are only ever
retained as inert blobs when exposed or preserved. Litchi never executes or
activates them.

## Focused Keynote transition edits

Enable the `keynote` feature to use the focused Keynote package API through
`litchi::keynote`. Transitions are selected by semantic slide name or position;
native identifiers are never exposed. Only an existing modern transition
envelope is editable, and `clear` is idempotent while retaining Keynote's
native no-effect envelope.

```toml
[dependencies]
litchi = { version = "0", default-features = false, features = ["keynote"] }
```

```rust,no_run
use std::io;

use litchi::keynote::{
    Package, SlideSelector,
    transition::Effect,
};

let package = Package::open("input.key")?;
let selector = SlideSelector::name("Appendix");
let mut settings = package
    .slide_transition(selector)?
    .ok_or_else(|| io::Error::other("slide has no modern transition"))?;
settings.set_effect(Some(Effect::Dissolve))?;
let commit = package.edit_slide_transition(selector)?.set(settings)?.commit()?;

let restored = commit
    .package()
    .apply_slide_transition(&commit.patch().inverse())?;
let mut output = Vec::new();
restored.package().write_to(&mut output)?;
assert!(!output.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

For safe filesystem publication, use
`litchi-keynote/examples/edit_slide_transition.rs`, which writes a distinct
output through `Package::write_to` into a synchronized sibling temporary file
and publishes it without clobbering an existing target.

## Focused Keynote title/body placeholder visibility

With the `keynote` feature enabled, `litchi::keynote` exposes the frozen
selector-first visibility transaction. `Kind::Title` and `Kind::Body` select
only existing layout-provided placeholders: a read returns `None` if the role
is missing, `Some(State::Visible)` if it draws, or `Some(State::Hidden)` if it
is retained without drawing. Hidden placeholders keep their text.

```rust,no_run
use litchi::keynote::{
    Package, SlideSelector,
    slide::placeholder::{Kind, State},
};

let package = Package::open("input.key")?;
let slide = SlideSelector::index(0);
let body_before = package.slide_body(slide)?;

let commit = package
    .edit_slide_placeholder_visibility(slide, Kind::Body)?
    .set(State::Hidden)
    .commit()?;
assert_eq!(
    commit
        .package()
        .slide_placeholder_visibility(slide, Kind::Body)?,
    Some(State::Hidden),
);
assert_eq!(commit.package().slide_body(slide)?, body_before);

let restored = commit
    .package()
    .apply_slide_placeholder_visibility(&commit.patch().inverse())?;
let mut output = Vec::new();
restored.package().write_to(&mut output)?;
assert!(!output.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

The transaction types are in `litchi::keynote::slide::placeholder`. Missing
roles return `Error::PlaceholderNotFound` from the edit entry point; visibility
does not create a placeholder. Slide-number visibility, layout edits, and
slide/placeholder creation are out of scope. An unchanged `set` is exact
no-op, and an inverse is exact-source checked. For safe publication, use
`litchi-keynote/examples/edit_slide_placeholder_visibility.rs`, which streams
through `Package::write_to` into a synchronized sibling temporary file and
publishes without clobbering an existing target.

## Focused Numbers table-header edits

Enable the `numbers` feature to use the focused immutable Numbers package API
through `litchi::numbers`. Header and footer settings use sheet and
sheet-scoped table selectors, never native object IDs. Assigning `None` clears
a field to native absence (effective zero or `false`); use `Some(Count::ONE)`
for a footer row and `Some(true)`/`Some(false)` to explicitly freeze or repeat
headers.

```toml
[dependencies]
litchi = { version = "0", default-features = false, features = ["numbers"] }
```

```rust,no_run
use litchi::numbers::{
    Package, SheetSelector, TableSelector,
    table::headers::{Count, Settings},
};

let package = Package::open("input.numbers")?;
let sheet = SheetSelector::name("Summary");
let table = TableSelector::name("Revenue");
let mut settings: Settings = package.table_header_settings(sheet, table)?;
settings.header_rows = Some(Count::TWO);
settings.footer_rows = Some(Count::ONE);
settings.header_rows_frozen = Some(true);
settings.repeating_header_rows_enabled = Some(true);

let commit = package
    .edit_table_headers(sheet, table)?
    .set(settings)
    .commit()?;

let restored = commit
    .package()
    .apply_table_headers(&commit.patch().inverse())?;
let mut output = Vec::new();
restored.package().write_to(&mut output)?;
assert!(!output.is_empty());
# Ok::<(), Box<dyn std::error::Error>>(())
```

The transaction types are in
`litchi::numbers::table::headers::transaction`. An unchanged edit is an exact
no-op; a changed edit fully reopens its candidate and removes stale root
previews. Applying the inverse to the committed package restores the exact
source. Changed legacy nested `Index.zip` input is refused with
`transaction::Error::UnsupportedSource` rather than being normalized. See
`litchi-numbers/examples/edit_table_headers.rs` for safe distinct-output,
no-clobber publication with `Package::write_to`.

Footer rows and row/column freeze flags are in the supported focused scope.
The effective boolean accessors are a native-bool oracle, reporting only the
optional native value (`None` is effectively `false`), never inferred UI state.
For safety, the transaction returns `UnsupportedDependency` for any
header/footer section-count change with an active pivot or group, header
row/column count changes backed by a rooted header-name manager, and repeat
flag changes on a legacy sheet topology.

## Feature Flags

Default features are empty. Enable only what the application needs; spelling
`default-features = false` is optional but valid.

```toml
# Legacy and OOXML PowerPoint, plus signing support.
litchi = { version = "0", features = ["ppt", "pptx", "sign"] }

# A minimal OOXML spreadsheet dependency.
litchi = { version = "0", default-features = false, features = ["xlsx"] }
```

Format leaves: `doc`, `docx`, `ppt`, `pptx`, `xls`, `xlsx`, `xlsb`, `rtf`,
`odt`, `ods`, `odp`, `pages`, `keynote`, and `numbers`.

Infrastructure: `cfb`, `ole`, `opc`, `ooxml-common`, `drawingml`,
`odf-common`, and `sheet`.

Capabilities: `sign`, `encryption`, `formula`, `fonts`, `images`, `eval`,
`web-functions`, `markdown`, and `yaml`.

`fonts` enables automatic system-font discovery and the shared
`litchi::fonts::embedding::Mode` publication policy. It forwards font embedding
to whichever of the independent `docx` and `pptx` leaves are enabled; enable a
format leaf alongside `fonts` to author a package with embedded fonts.

Convenience aggregates: `legacy`, `ooxml`, `odf`, `iwork`, `word`, `slides`,
`spreadsheets`, `office`, `all-formats`, and `all`.

`pages`, `keynote`, and `numbers` are independent full parsing leaves. Their
concrete owner modules are `litchi::pages`, `litchi::keynote`, and
`litchi::numbers`; `iwork` enables all three without adding another API layer.

Formats do not implicitly enable signing; add `sign` explicitly when needed.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
