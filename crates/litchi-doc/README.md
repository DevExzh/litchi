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
source text. `body_text::Snapshot::paragraphs` additionally exposes stored,
accepted, and rejected ordinary body-paragraph projections. Its focused edit
seam supports bounded length-changing replacements across multiple ordinary
paragraphs plus direct-bold changes in one immutable transaction. It rebuilds
the CLX and CHPX FKPs, shifts its modeled main-story PLCFs, fully reopens the
result, and explicitly refuses structural, tracked, mixed-format, or unmodeled
position dependencies. Source-checked in-memory and durable semantic patches,
disjoint composition, and bounded undo/redo history use the same operation
model.
Paragraph selection uses the format-neutral, zero-based `litchi_core::Position`;
resolving it against a source body reports a typed not-found refusal.

Format-neutral OfficeArt image discovery lives in `litchi-odraw`; optional
codec operations are provided by the separate `litchi-imgconv::Convert`
extension trait.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
