# litchi-odf

OpenDocument Format (ODF) umbrella for the independently selectable family
crates covering `.odt`, `.ods`, `.odp`, `.odg`, `.odc`, `.odi`, `.odm`, `.oth`,
and `.odb` files.

## Overview

`litchi-odf` provides detection and an optional facade over independently
selectable family crates. For the smallest dependency and memory footprint,
depend directly on the family crate you need. Each family keeps its API
contextual and layered: for example, `litchi_odt::Document`,
`litchi_ods::names::{Definition, Range}`, and `litchi_odp::slide::Slide`.

## Usage

```toml
[dependencies]
litchi-odf = "0.0.1"
```

```rust
use litchi_odf::{odp, ods, odt};

// Each family's immutable `Package` owns every state verb; the
// `Document`, `Workbook` and `Presentation` projections are read-only.
let package = odt::Package::open("document.odt")?;
let text = package.document()?.text()?;

let sheet = ods::Package::open("data.ods")?.workbook()?;
let names = sheet.sheet_names();

let pres = odp::Package::open("slides.odp")?.presentation()?;
let slide_values: Vec<odp::slide::Slide> = pres.slides()?;
let slide_count = pres.slide_count()?;

// Edits stage on a transaction and publish a new package; `save` is an
// atomic sibling-temporary replacement.
let mut edit = package.edit();
edit.replace_paragraph(odt::transaction::Position::new(0), "Updated")?;
edit.commit()?.package().save("updated.odt")?;
# Ok::<(), litchi_core::Error>(())
```

The umbrella root intentionally exposes only detection and the selected family
modules. Shared vocabulary and package primitives remain under the separate
`litchi-odf-common` crate. The default umbrella features are `odt`, `ods`, and
`odp`; `--no-default-features` enables detection only, while `all` opts into
every family module.

## Features

- Read and write ODF text documents (`.odt`) with paragraphs, lists, tables, and styles
- Read and write ODF spreadsheets (`.ods`) with typed cell values and formulas
- Read and write ODF presentations (`.odp`) with slides and shapes
- Read and write standalone drawings (`.odg`), charts (`.odc`), images (`.odi`),
  master documents (`.odm`), web templates (`.oth`), and database front ends
  (`.odb`) through their optional family facades
- Metadata extraction (title, author, statistics) through the selected family
  modules

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
