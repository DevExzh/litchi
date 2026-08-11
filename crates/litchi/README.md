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

## Focused Pages section-settings edits

Enable the `pages` feature to use the selector-first Pages package API through
`litchi::pages`. `section::Settings` is the complete aggregate value for this
transaction: its optional name, four optional Boolean flags, and three optional
pagination fields retain native presence. In particular, `None` removes a
Boolean field while `Some(false)` preserves an explicitly encoded false.
Transaction, patch, diagnostics, error, limit, dependency, and path types live
under `litchi::pages::section::settings`; none exposes a native identifier.

```toml
[dependencies]
litchi = { version = "0", default-features = false, features = ["pages"] }
```

```rust,no_run
use litchi::pages::{Package, SectionSelector, section::Settings};

let package = Package::open("input.pages")?;
let selector = SectionSelector::name("Introduction");
let mut settings: Settings = package.section_settings(selector)?;
settings.set_inherit_previous_header_footer(Some(false));
settings.set_first_page_hides_header_footer(Some(true));

let commit = package
    .edit_section_settings(selector)?
    .set(settings)?
    .commit()?;
assert_eq!(
    &commit.package().section_settings(selector)?,
    commit.patch().after(),
);

let restored = commit
    .package()
    .apply_section_settings(&commit.patch().inverse())?;
assert_eq!(restored.package().source_bytes(), package.source_bytes());
# Ok::<(), Box<dyn std::error::Error>>(())
```

An unchanged replacement is an exact no-op and retains the source allocation,
layout cache, and previews. A changed exact-source transaction proves the
selected section's template dependencies, patches only the selected Section
component, preserves ViewState, the rooted derived layout cache, and canonical
root previews exactly, and fully reopens the candidate before publication. Its
inverse is authorized only against the exact committed artifact and restores
the exact source. Changed legacy nested packages are refused instead of
normalized.

Pages 14.4 opened, saved, closed, and exact-path reopened matched Rust
artifacts without warnings. Separate pairs proved only explicit-false-to-true
changes for header/footer inheritance, even/odd distinction, and first-page
header/footer hiding; the section header, templates, and all 36 text storages
remained exact. First-page-different stayed explicitly false, and the UI does
not prove absent-versus-false behavior. Native previews also stayed exact, so
the focused writer retains them together with the layout cache and all other
ViewState bytes. A changed native-supported edit touches one Section component
and reports zero deleted previews.

The focused section-name and section-pagination APIs remain ergonomic facades
for their individual value families. Use the aggregate transaction when the
four flags, name, or pagination must change atomically. For safe filesystem
publication, `litchi-pages/examples/edit_section_settings.rs` preserves the
source name and pagination while changing the four Boolean fields, writes a
distinct output through a synchronized sibling temporary file, and never
clobbers an existing target.

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

## Focused Keynote per-slide slide-number visibility

`Kind::SlideNumber` uses the same `litchi::keynote` selector-first transaction
for one existing slide. It is distinct from the presentation-wide
`show::Settings::slide_numbers_visible` flag: use the show-settings transaction
when changing the whole presentation. A per-slide read returns `None` if the
layout provides no slide-number placeholder, `Some(State::Visible)` if it
draws, and `Some(State::Hidden)` when its placeholder, storage/text graph, and
layout reference are retained without per-slide drawing.

```rust,no_run
use litchi::keynote::{
    Package, SlideSelector,
    slide::placeholder::{Kind, State},
};

let package = Package::open("input.key")?;
let slide = SlideSelector::index(0);
let commit = package
    .edit_slide_placeholder_visibility(slide, Kind::SlideNumber)?
    .set(State::Hidden)
    .commit()?;
assert_eq!(
    commit
        .package()
        .slide_placeholder_visibility(slide, Kind::SlideNumber)?,
    Some(State::Hidden),
);

let restored = commit
    .package()
    .apply_slide_placeholder_visibility(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

This transaction does not create a placeholder or modify its storage/text,
layout, or slide creation policy. It is an exact no-op for an unchanged state;
a changed exact-source commit invalidates stale rendering state and its inverse
restores the exact source. Use
`litchi-keynote/examples/edit_slide_number_visibility.rs` for distinct-output,
sibling-temporary, no-clobber publication with `Package::write_to`.

## Focused Keynote soundtrack playback edits

With the `keynote` feature enabled, `litchi::keynote::soundtrack` provides the
immutable playback-settings transaction for an existing soundtrack. `None`
from `soundtrack_settings()` means the soundtrack object is absent, while
`Some(Settings::default())` means it exists with both optional values absent.
The transaction exposes semantic mode and volume settings with direct `Edit`,
`Patch`, `Commit`, diagnostics, error, and limit types; native identifiers are
never exposed.

```rust,no_run
use std::io;

use litchi::keynote::{
    Package,
    soundtrack::{Mode, Settings},
};

let package = Package::open("input.key")?;
let before = package
    .soundtrack_settings()?
    .ok_or_else(|| io::Error::other("presentation has no soundtrack"))?;
let mut settings: Settings = before;
settings.set_volume(Some(0.35))?;
settings.set_mode(Some(Mode::Loop))?;
let commit = package.edit_soundtrack_settings()?.set(settings).commit()?;
assert!(commit.diagnostics().changed());
assert_eq!(commit.diagnostics().touched_components(), 1);

let restored = commit
    .package()
    .apply_soundtrack_settings(&commit.patch().inverse())?;
assert_eq!(restored.package().soundtrack_settings()?, Some(before));
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

Passing `None` to either settings setter clears native presence, and
`Mode::Unknown` retains future native values (known values use named variants).
An unchanged edit is exact no-op. A changed edit rewrites only the owning
soundtrack component, fully reopens its candidate without invalidating
rendering previews, and its inverse applies only to the exact committed
snapshot. Media entries, assets, ordering, and references remain in the legacy
migration host's retained collection scope. Use
`litchi-keynote/examples/edit_soundtrack_settings.rs` for synchronized,
distinct-output, no-clobber publication through `Package::write_to`.

## Focused Numbers sheet-order edits

Enable the `numbers` feature to reorder existing sheets through
`litchi::numbers::sheet::order`. A sheet selector uses its exact name or
zero-based position; the destination is the moved sheet's final position. The
transaction updates the document and navigator/sidebar orders together without
exposing physical identifiers.

```rust,no_run
use litchi::numbers::{Package, SheetSelector};

let package = Package::open("input.numbers")?;
let commit = package
    .edit_sheet_order()
    .move_sheet(SheetSelector::name("Archive"), 0)?
    .commit()?;
assert!(commit.diagnostics().changed());
assert_eq!(commit.diagnostics().touched_components(), 1);
assert_eq!(commit.diagnostics().deleted_previews(), 3);
assert!(commit.diagnostics().full_reparse_performed());

let restored = commit
    .package()
    .apply_sheet_order(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

A move to the current position is exact no-op. A changed edit updates one
owning component, removes the three stale root previews, and fully reopens its
candidate; its inverse restores all three previews and the exact source. The
native gate verifies a warning-free Numbers open plus Save As/reopen. For safe
publication, use `litchi-numbers/examples/edit_sheet_order.rs`, which writes a
distinct output via `Package::write_to`, a synchronized sibling temporary
file, and no-clobber publication.

## Focused Numbers table-title edits

With the `numbers` feature enabled, table-title visibility and outline use the
selector-first `litchi::numbers::table::title` transaction. `Settings` keeps
native presence: `None` means absent, while `Some(false)` explicitly stores a
false value. Explicit false and outline presence are losslessly transaction
tested, but not native UI-oracle claims. Native table IDs are never exposed.

```rust,no_run
use litchi::numbers::{
    Package, SheetSelector, TableSelector,
    table::title::Settings,
};

let package = Package::open("input.numbers")?;
let sheet = SheetSelector::name("Summary");
let table = TableSelector::name("Revenue");
let before = package.table_title_settings(sheet, table)?;
// This is guaranteed to differ without enabling a previously hidden title.
let settings = if before.visible() == Some(true) {
    Settings::new(None, before.outlined())
} else {
    Settings::new(
        before.visible(),
        match before.outlined() {
            None => Some(false),
            Some(_) => None,
        },
    )
};
let commit = package
    .edit_table_title(sheet, table)?
    .set(settings)
    .commit()?;
assert_eq!(commit.package().table_title_settings(sheet, table)?, settings);
assert!(commit.diagnostics().changed());
assert_eq!(commit.diagnostics().touched_components(), 1);
assert!(commit.diagnostics().deleted_previews() <= 3);

let restored = commit
    .package()
    .apply_table_title(&commit.patch().inverse())?;
let mut original = Vec::new();
package.write_to(&mut original)?;
let mut restored_bytes = Vec::new();
restored.package().write_to(&mut restored_bytes)?;
assert_eq!(restored_bytes, original);
# Ok::<(), Box<dyn std::error::Error>>(())
```

An unchanged value is exact no-op. A changed title updates its one
`CalculationEngine` component, removes every existing canonical root preview,
and fully reopens the candidate; its inverse restores the exact source and
previews. Changed publication refuses an effectively locked table, so the
changed portion of this example requires an unlocked supported table. A visible
title also requires the native title-height, paragraph-style, and shape-style
prerequisites. The native basic fixture has all three canonical previews
(`3 → 0`, then `0 → 3` on inverse) and proves only visible `Some(true)` to
absent (hide), warning-free open, and Save As/reopen; it does not assert
outline or explicit-false UI results. Use
`litchi-numbers/examples/edit_table_title.rs` for synchronized,
distinct-output, no-clobber publication via `Package::write_to`.

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
