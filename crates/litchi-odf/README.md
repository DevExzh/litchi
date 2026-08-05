# litchi-odf

OpenDocument Format (ODF) umbrella for the independently selectable family
crates covering `.odt`, `.ods`, `.odp`, `.odg`, `.odc`, `.odi`, `.odm`, `.oth`,
and `.odb` files.

## Overview

`litchi-odf` provides detection, common ODF vocabulary, and an optional facade
over independently selectable family crates. For the smallest dependency and
memory footprint, depend directly on the family crate you need. Each family
keeps its API contextual: for example, `litchi_odg::Builder` and
`litchi_oth::{Builder, Template}`.

## Usage

```toml
[dependencies]
litchi-odf = "0.0.1"
```

```rust
use litchi_odf::{odp, ods, odt};

let doc = odt::Document::open("document.odt")?;
let text = doc.text()?;

let sheet = ods::Spreadsheet::open("data.ods")?;
let csv = sheet.to_csv()?;

let pres = odp::Presentation::open("slides.odp")?;
let slides = pres.slide_count()?;
# Ok::<(), litchi_core::Error>(())
```

The umbrella root intentionally exposes only detection and the selected family
modules. Shared vocabulary and package primitives remain under the separate
`litchi-odf-common` crate, so applications that need only one family can keep
their dependency and memory footprint minimal.

## Features

- Read and write ODF text documents (`.odt`) with paragraphs, lists, tables, and styles
- Read and write ODF spreadsheets (`.ods`) with typed cell values and formulas
- Read and write ODF presentations (`.odp`) with slides and shapes
- Read and write standalone drawings (`.odg`), charts (`.odc`), images (`.odi`),
  master documents (`.odm`), web templates (`.oth`), and database front ends
  (`.odb`) through their optional family facades
- Metadata extraction (title, author, statistics) for all three formats

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
