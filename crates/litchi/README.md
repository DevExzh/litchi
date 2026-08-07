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
