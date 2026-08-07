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

## Bounded Ingestion

All ordinary package constructors use bounded `ReadLimits::default()` values.
For untrusted or multi-tenant input, build a checked profile from those defaults
and use an `*_with_limits` constructor:

```rust
use litchi_opc::{OpcPackage, ReadLimits};

let limits = ReadLimits::builder()
    .max_input_bytes(32 * 1024 * 1024)?
    .max_archive_members(10_000)?
    .build()?;
let package = OpcPackage::open_with_limits("untrusted.docx", limits)?;
# let _ = package;
# Ok::<(), litchi_opc::OpcError>(())
```

The standard profile caps input at 512 MiB; ZIP members and materialized OPC
parts at 100,000 each; one ZIP member and one materialized part at 512 MiB;
and aggregate declared ZIP bytes at 2 GiB. It also bounds ZIP names, central
directory metadata, compressed bytes, aggregate materialized parts,
`[Content_Types].xml` and its mappings, relationship parts and XML, individual
and aggregate relationships, graph nodes, XML events and depth, and
relationship attribute and target lengths. `ReadResource` identifies the
specific rejected resource.

These ceilings are Litchi safety policy rather than ECMA-376 size maxima. They
implement a defensive consumer boundary for the physical package and
relationships described by ECMA-376 Part 2 sections 7.3.6 and 10, and the
corresponding MS-OI29500 sections 2.1.1749-1752. DOCX and PPTX expose
`Package::*_with_limits`; XLSX exposes both `Package::*_with_limits` and
`Workbook::*_with_limits`; XLSB exposes `Workbook::new_with_limits`. Each takes
the same `ReadLimits` profile so callers can use one contextual policy across
OOXML formats.

OPC readers treat macros, VBA, ActiveX, controls, OLE objects, and embedded
code as inert blobs only when they are retained or exposed. They are never
executed or activated.

## Features

- Streaming OPC package reader with content-type and relationship resolution
- `PackURI` parsing and normalisation per ISO/IEC 29500-2
- Zero-copy XML parsing via `quick-xml`, SIMD integer parsing via `atoi_simd`
- `PackageWriter` for authoring new OPC packages

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
