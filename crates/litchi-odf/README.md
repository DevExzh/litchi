# litchi-odf

OpenDocument Format (ODF) reader and writer for `.odt`, `.ods`, and `.odp` files.

## Overview

`litchi-odf` parses and produces OpenDocument files conforming to ISO/IEC 26300 (ODF 1.2): text documents (`.odt`), spreadsheets (`.ods`), and presentations (`.odp`). It includes a ZIP-based package layer, XML element model for content / styles / metadata, and high-level types for each document family. Builders and `Mutable*` views are provided for creating and editing files in addition to read-only access.

## Usage

```toml
[dependencies]
litchi-odf = "0.0.1"
```

```rust
use litchi_odf::{Document, Spreadsheet, Presentation};

let mut doc = Document::open("document.odt")?;
let text = doc.text()?;

let mut sheet = Spreadsheet::open("data.ods")?;
let csv = sheet.to_csv()?;

let mut pres = Presentation::open("slides.odp")?;
let slides = pres.slide_count()?;
# Ok::<(), litchi_core::Error>(())
```

## Features

- Read and write ODF text documents (`.odt`) with paragraphs, lists, tables, and styles
- Read and write ODF spreadsheets (`.ods`) with typed cell values and formulas
- Read and write ODF presentations (`.odp`) with slides and shapes
- Metadata extraction (title, author, statistics) for all three formats

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
