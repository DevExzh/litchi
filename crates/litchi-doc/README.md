# litchi-doc

Strictly typed support for the legacy Microsoft Word (`.doc`) binary format.

## Overview

This crate owns legacy Word parsing and writing. It builds on `litchi-cfb` for
the CFB storage substrate, `litchi-ole-common` for shared OLE object metadata,
and `litchi-odraw` for OfficeArt. Its canonical API lives at the crate root;
there is no migration-host or format-dispatch layer.

## Usage

```toml
[dependencies]
litchi-doc = "0.0.1"
```

```rust
use litchi_doc::{Limits, Package, PackageOpenOptions};

let limits = Limits::default().with_max_package_bytes(32 * 1024 * 1024)?;
let mut package = Package::open_with("example.doc", PackageOpenOptions::default().with_limits(limits))?;
let document = package.document()?;
println!("paragraphs: {}", document.paragraph_count()?);
# Ok::<(), litchi_doc::Error>(())
```

## Features

- `.doc` (Word 97-2003) reader and writer with full PLCF/SPRM handling
- Optional `formula` feature for MathType (MTEF) extraction

`Package::open` uses finite defaults (128 MiB package, 64 MiB stream, and
96 MiB aggregate DOC streams). Use `Package::open_with` and
`PackageOpenOptions` to select stricter or workload-specific limits. Passwords
are supplied to encrypted documents through `OpenOptions::with_password` and
the non-cloneable, zeroizing `Password` type. `Document::text` exposes stored
source text. `body_text::Snapshot` additionally exposes stored, accepted, and
rejected body projections plus paragraph selectors for every DOC story, simple
table cells, simple cached field results, and revision marks. Its bounded
immutable transaction can resize modeled main-story text, overwrite
equal-length Unicode text in auxiliary stories, apply direct
bold/italic/underline, dispose of non-destructive revision marks, and change
passive embedded-object display metadata. It rebuilds CLX and CHPX FKPs, shifts
the modeled CP/PLCF closure, updates FIB story counts, and fully reopens the CFB
and DOC before publication. Source-checked in-memory and durable semantic
patches, disjoint composition, three-way planning, dependency-free text
transfer, and bounded undo/redo history use the same operation model.
Structural table edits, field delimiters or nesting, destructive revision
dispositions, auxiliary-story length changes, mixed formatting, and unmodeled
CP dependencies are typed refusals. OfficeArt, drawing payload, and resource
transfer remain with their dedicated owners and are not silently copied
through this text transaction. Selection uses the format-neutral, zero-based
`litchi_core::Position`; resolving it against a source collection reports a
typed not-found refusal.

Format-neutral OfficeArt image discovery lives in `litchi-odraw`; optional
codec operations are provided by the separate `litchi-imgconv::Convert`
extension trait.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
