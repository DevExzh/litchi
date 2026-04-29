# litchi-ooxml

Office Open XML (OOXML) reader and writer for `.docx`, `.xlsx`, `.xlsb`, and `.pptx` files.

## Overview

`litchi-ooxml` implements the Office Open XML formats defined by ECMA-376 / ISO/IEC 29500 for modern Microsoft Office documents. It builds on top of [`litchi-opc`](../litchi-opc), which provides the underlying Open Packaging Conventions (ZIP container, parts, and relationships) layer. Format-specific submodules (`docx`, `xlsx`, `xlsb`, `pptx`) expose package types for opening, reading, and writing each document family.

## Usage

```toml
[dependencies]
litchi-ooxml = "0.0.1"
```

```rust
use litchi_ooxml::docx::Package;

let pkg = Package::open("document.docx")?;
let doc = pkg.document()?;
let text = doc.text()?;
println!("paragraphs: {}", doc.paragraph_count()?);
# Ok::<(), Box<dyn std::error::Error>>(())
```

## Features

- Read and write Word documents (`.docx`)
- Read and write Excel workbooks (`.xlsx` and binary `.xlsb`)
- Read and write PowerPoint presentations (`.pptx`)
- Optional `encryption` feature for password-protected OOXML packages
- Optional `fonts` feature for embedded font handling

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
