# litchi-rtf

Parser and writer for the Rich Text Format (RTF), targeting the RTF 1.9.1 specification.

## Overview

This crate provides a bounded RTF reader and streaming writer with an
immutable, cheap-to-share `Document` facade. Borrowed semantic views traverse
paragraphs, runs, and structural breaks without first flattening the retained
document. Borrowed font catalogs and color palettes resolve run formatting
without exposing numeric RTF table IDs. The advanced retained model covers
tables, lists, sections, fields, pictures, shapes, and stylesheets. The crate
also handles the compressed RTF transport used inside MAPI messages.

## Usage

```toml
[dependencies]
litchi-rtf = "0.0.1"
```

```rust
use litchi_rtf::Document;

let rtf = r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello World!\par}";
let doc = Document::parse(rtf)?;
assert_eq!(doc.text(), "Hello World!\n");

for paragraph in doc.body().paragraphs() {
    println!("{paragraph}");
}

let first_run = doc.body().runs().next().expect("body text");
assert_eq!(
    first_run.format().font().map(|font| font.name()),
    Some("Helvetica")
);
# Ok::<(), litchi_rtf::Error>(())
```

## Features

- Lexer + parser covering control words, groups, and binary data
- Document model: paragraphs, runs, tables, lists, sections, fields, pictures, shapes
- Stylesheet, font table, and color table handling
- Compressed RTF (`MS-OXRTFCP`) encode/decode
- Immutable, cheap-to-share `Document` snapshots for ordinary reads
- Bounded source-checked `Document::edit()` transactions composing disjoint
  UTF-8 body spans, paragraph alignment, character bold ranges, ordinary
  paragraph insertion, table-cell text, and header/footer text, with atomic
  commits, reversible durable patches, deterministic sub-edit/three-way
  composition and commit-coupled bounded history
- Checked ordinary-root transfer plans for passive fields, complete nested
  table trees, style dependency closures, lists with overrides, and inert
  embedded objects with remapped result pictures; opaque target destinations,
  active links, and unresolved resource collisions fail closed
- Lazy borrowed `text::Story`, paragraph, inline, and run traversal
- Sparse-safe `font::Catalog` and checked `color::Palette` resource views
- Semantic run font/color resolution without numeric table references
- Distinct semantic paragraph (`\\par`) and line (`\\line`) boundaries
- Concise `read`, `write`, and `transport` modules for format operations
- Streaming `write::Writer` with configurable `write::Options`

## Source organization

The implementation is grouped by responsibility rather than as a flat module
list:

- `api`: the immutable facade and borrowed semantic story views
- `codec`: compressed transport, limits, lexer, parser, and writer
- `model`: retained document storage and native value types
- `resource`: borrowed font catalogs and color palettes
- `text`, `content`, and `drawing`: authored document content
- `review`, `metadata`, `numbering`, and `policy`: supporting document state

These directories are real internal module boundaries rather than file-only
groupings. Existing public modules such as `text`, `field`, `table`, `picture`,
and `review` remain stable through contextual re-exports.

The attached mutable retained tree is isolated under the explicit advanced
`raw` module. Ordinary application code reads immutable `Document` snapshots
and publishes changes only through `Document::edit()` or composed sub-edits.
Canonical edits of retained destinations refuse snapshots containing unknown
syntax; exact body splices preserve unknown destinations byte-for-byte.

`transport::decompress` enforces a finite 256 MiB expansion ceiling before
allocation. Applications with a different document budget can call
`transport::decompress_with_limits` with an explicit `transport::Limits`
value. Document parsing and file opening also use finite source, token,
binary-payload, and expansion ceilings through `read::Limits`; custom profiles
are accepted by `Document::parse_with_limits`, `Document::from_bytes_with_limits`,
and `Document::open_with_limits`.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
