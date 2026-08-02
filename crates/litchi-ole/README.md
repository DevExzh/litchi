# litchi-ole

Migration host for the legacy Microsoft Word (`.doc`) binary format.

## Overview

This crate parses and writes the remaining OLE2-based Word format while its
concrete `litchi-doc` crate is extracted. PowerPoint ownership has moved to
`litchi-ppt`, and Excel BIFF ownership has moved to `litchi-xls`. This crate
intentionally provides neither `ppt`/`xls` modules nor compatibility
re-exports. It builds on `litchi-cfb` for the CFB storage substrate and retains
only DOC migration-host infrastructure.

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
- Optional `formula` feature for MathType (MTEF) extraction

Format-neutral OfficeArt image discovery lives in `litchi-odraw`; optional
codec operations are provided by the separate `litchi-imgconv::Convert`
extension trait.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
