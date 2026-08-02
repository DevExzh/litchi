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
use litchi_rtf::RtfDocument;

let rtf = r"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard Hello World!\par}";
let doc = RtfDocument::parse(rtf)?;
let text = doc.text();
# Ok::<(), litchi_rtf::RtfError>(())
```

## Features

- Lexer + parser covering control words, groups, and binary data
- Document model: paragraphs, runs, tables, lists, sections, fields, pictures, shapes
- Stylesheet, font table, and color table handling
- Compressed RTF (`MS-OXRTFCP`) encode/decode
- `RtfWriter` with configurable `WriterOptions` for round-tripping documents

`decompress` enforces a finite 256 MiB expansion ceiling before allocation.
Applications with a different document budget can call `decompress_with_limits`
with a checked `DecompressionLimits` value.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
