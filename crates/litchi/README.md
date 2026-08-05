# litchi

High-performance Rust library for parsing Microsoft Office, OpenDocument, and Apple iWork file formats with a unified API.

## Overview

`litchi` is the user-facing umbrella crate of the [Litchi workspace](https://github.com/DevExzh/litchi). It auto-detects file formats and delegates parsing to independently owned format crates (`litchi-doc`, `litchi-ppt`, `litchi-xls`, `litchi-docx`, `litchi-pptx`, `litchi-xlsb`, `litchi-xlsx`, `litchi-opc`, `litchi-ooxml-common`, `litchi-odf`, `litchi-iwa`, `litchi-rtf`, and friends). Most users should depend on this crate rather than the format-specific ones. Canonical low-level legacy-format entry points are `litchi::doc`, `litchi::ppt`, and `litchi::xls`; there is no compatibility facade nesting them under an OLE module.

Shared OOXML chart and SmartArt grammar is available through the concise
`litchi::drawing::{chart, diagram}` facade when the `ooxml` feature is enabled.
Concrete formats retain their package-specific anchors and relationships.

## Usage

```toml
[dependencies]
litchi = "0.0.1"
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

Default: `doc`, `ppt`, `xls`, `ooxml`, `ooxml_encryption`, `eval_engine`.

| Flag | Adds support for |
|------|------------------|
| `doc` | Legacy Word `.doc` |
| `ppt` | Legacy PowerPoint `.ppt` |
| `xls` | Legacy Excel BIFF `.xls` |
| `ooxml` | `.docx`, `.xlsx`, `.xlsb`, `.pptx` |
| `ooxml_encryption` | Password-protected OOXML files |
| `odf` | `.odt`, `.ods`, `.odp` |
| `iwa` | Apple `.pages`, `.numbers`, `.key` |
| `rtf` | Rich Text Format |
| `formula` | MathType / OMML to LaTeX conversion |
| `imgconv` | EMF / WMF / PICT image conversion |
| `fonts` | Font discovery and subsetting |
| `eval_engine` | Spreadsheet formula evaluation (`=SUM(A1:A10)`) |
| `markdown` | Markdown emission helpers |
| `full` | Enables every feature above |

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
