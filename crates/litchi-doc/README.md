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
use litchi_doc::Package;

let mut package = Package::open("example.doc")?;
let document = package.document()?;
println!("paragraphs: {}", document.paragraph_count()?);
# Ok::<(), litchi_doc::DocError>(())
```

## Features

- `.doc` (Word 97-2003) reader and writer with full PLCF/SPRM handling
- Optional `formula` feature for MathType (MTEF) extraction

Format-neutral OfficeArt image discovery lives in `litchi-odraw`; optional
codec operations are provided by the separate `litchi-imgconv::Convert`
extension trait.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
