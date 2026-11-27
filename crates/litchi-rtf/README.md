# litchi-rtf

Parser and writer for the Rich Text Format (RTF), targeting the RTF 1.9.1 specification.

## Overview

This crate provides a bounded RTF reader and streaming writer with an
immutable, cheap-to-share `Package` owner. `Package::document()` returns an
owned read-only projection for semantic traversal. Borrowed views traverse
paragraphs, runs, and structural breaks without first flattening the retained
package. Borrowed font catalogs and color palettes resolve run formatting
without exposing numeric RTF table IDs. The advanced retained model covers
tables, lists, sections, fields, pictures, shapes, stylesheets, positioned
paragraph frames, inert SmartTag/factoid and move-bookmark ranges, and modern
protection password-hash records. The crate
also handles the compressed RTF transport used inside MAPI messages.

## Usage

```toml
[dependencies]
litchi-rtf = "0.0.1"
```

```rust
use litchi_rtf::Package;

let rtf = r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello World!\par}";
let package = Package::parse(rtf)?;
let doc = package.document();
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

See the [feature matrix](docs/FEATURE_MATRIX.md) for read/write coverage,
ordinary transaction limits, retained-model support, and explicit semantic gaps.

- Lexer + parser covering control words, groups, and binary data
- Document model: paragraphs, runs, tables, lists, sections, fields, pictures, shapes
- Typed positioned paragraph frames, bounded inert SmartTag/factoid and
  move-bookmark metadata, and bounded inert `\\passwordhash` records
- Stylesheet, font table, and color table handling
- Compressed RTF (`MS-OXRTFCP`) encode/decode
- Immutable, cheap-to-share `Package` owners with read-only `Package::document()` projections
- Bounded source-checked `Package::edit()` transactions composing disjoint
  UTF-8 body spans, paragraph alignment, character bold ranges, ordinary
  character italic ranges,
  paragraph insertion, table-cell text, header/footer text, comment bodies,
  footnote/endnote stories, and inert root shape text frames, with atomic
  commits, reversible durable patches, deterministic sub-edit/three-way
  composition and commit-coupled bounded history
- Checked ordinary-root transfer plans for passive fields, nested table trees
  with inert drawings/breaks/cell fields, style dependency closures,
  picture-bullet lists with remapped shared pictures and overrides, and inert
  embedded objects with remapped result pictures and root text-frame shapes;
  opaque target destinations, active links, and unresolved resource collisions
  fail closed
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
`raw` module. Ordinary application code opens immutable `Package` owners,
reads through the package or a `Package::document()` projection, and publishes
changes through `Package::edit()` or composed sub-edits. A projection shares
the package's immutable state and remains valid when its owner is dropped;
projections do not own lifecycle or publication operations.
Canonical edits of retained destinations refuse packages containing unknown
syntax; exact body splices preserve unknown destinations byte-for-byte.

`transport::{compress, decompress}` require explicit transport limit profiles:
`compress` takes `transport::CompressionLimits`, and `decompress` takes
`transport::DecompressionLimits`. `CompressionLimits::default()` and
`DecompressionLimits::default()` select a 256 MiB output ceiling, while
`CompressionLimits::new(max_output_bytes)` and
`DecompressionLimits::new(max_output_bytes)` let the caller choose another
finite bound; the selected output bound is checked before every fallible
growth and final frame allocation.
Package parsing and file opening use checked `SourceLimits` profiles
for input, events, binary payloads, expansion and opaque retention. Build a
profile with `SourceLimits::builder()` and fallible `max_*` setters, then pass
`OpenOptions::new().source_limits(profile)` to `Package::parse_with_options`,
`Package::from_bytes_with_options` or `Package::open_with_options`. Input
defaults to 512 MiB; zero explicitly denies a dimension. Explicit repair is
selected with `OpenOptions::leniency`.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
