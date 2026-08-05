# litchi-opc

Open Packaging Conventions (OPC) implementation: the ZIP-based container layer used by all OOXML formats.

## Overview

OPC defines how `.docx`, `.xlsx`, `.pptx`, and other Office Open XML documents
are packaged: parts, content types, and relationships inside a ZIP archive.
This crate provides the package model, `PackURI` resolution, and reader/writer
plumbing on top of `soapberry-zip` and `quick-xml`. It is consumed directly by
`litchi-docx`, `litchi-pptx`, `litchi-xlsx`, and `litchi-xlsb`; shared OOXML
vocabulary and graph services live in `litchi-ooxml-common`.

## Usage

```toml
[dependencies]
litchi-opc = "0.0.1"
```

```rust
use litchi_opc::{OpcPackage, Part};

let pkg = OpcPackage::open("example.docx")?;
for part in pkg.iter_parts() {
    println!("{} ({})", part.partname(), part.content_type());
}
# Ok::<(), litchi_opc::OpcError>(())
```

## Features

- Streaming OPC package reader with content-type and relationship resolution
- `PackURI` parsing and normalisation per ISO/IEC 29500-2
- Zero-copy XML parsing via `quick-xml`, SIMD integer parsing via `atoi_simd`
- `PackageWriter` for authoring new OPC packages

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
