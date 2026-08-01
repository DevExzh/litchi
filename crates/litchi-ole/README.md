# litchi-ole

Migration host for the legacy Microsoft Word (`.doc`) and PowerPoint (`.ppt`)
binary formats.

## Overview

This crate parses and writes the remaining OLE2-based Word and PowerPoint
formats while their concrete crates are extracted. Excel BIFF ownership has
moved to `litchi-xls`; this crate intentionally provides no `xls` module or
compatibility re-export. It builds on `litchi-cfb` for the CFB storage
substrate and retains only the DOC/PPT migration-host infrastructure.

## Usage

```toml
[dependencies]
litchi-ole = "0.0.1"
```

```rust
use litchi_ole::doc::Package;

let mut package = Package::open("example.doc")?;
let document = package.document()?;
println!("paragraphs: {}", document.paragraph_count()?);
# Ok::<(), litchi_ole::doc::DocError>(())
```

## Features

- `.doc` (Word 97-2003) reader and writer with full PLCF/SPRM handling
- `.ppt` (PowerPoint 97-2003) reader and writer with Escher drawings
- Optional `formula` feature for MathType (MTEF) extraction
- Optional `imgconv` feature for EMF/WMF/PICT image bridges

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
