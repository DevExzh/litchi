# litchi-rtf

Parser and writer for the Rich Text Format (RTF), targeting the RTF 1.9.1 specification.

## Overview

This crate provides a high-performance RTF reader and writer with arena
allocation (via `bumpalo`), zero-copy patterns where practical, and a
structured document model covering paragraphs, runs, tables, lists,
sections, fields, pictures, shapes, and stylesheets. It also handles the
compressed RTF transport (`is_compressed_rtf`, `compress`, `decompress`)
used inside MAPI messages.

## Usage

```toml
[dependencies]
litchi-rtf = "0.0.1"
```

```rust
use litchi_rtf::Document;

let rtf = r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello World!\par}";
let doc = Document::parse(rtf)?;
let text = doc.text();
# Ok::<(), litchi_rtf::Error>(())
```

## Features

- Lexer + parser covering control words, groups, and binary data
- Document model: paragraphs, runs, tables, lists, sections, fields, pictures, shapes
- Stylesheet, font table, and color table handling
- Compressed RTF (`MS-OXRTFCP`) encode/decode
- Immutable, cheap-to-share `Document` snapshots for ordinary reads
- Concise `read`, `write`, and `transport` modules for format operations
- Streaming `write::Writer` with configurable `write::Options`

`transport::decompress` enforces a finite 256 MiB expansion ceiling before
allocation. Applications with a different document budget can call
`transport::decompress_with_limits` with an explicit `transport::Limits`
value. Document parsing and file opening also use finite source, token,
binary-payload, and expansion ceilings through `read::Limits`; custom profiles
are accepted by `Document::parse_with_limits`, `Document::from_bytes_with_limits`,
and `Document::open_with_limits`.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
