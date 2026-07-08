# litchi-ole

Reader and writer for the legacy Microsoft Office binary formats: `.doc`, `.xls`, and `.ppt`.

## Overview

This crate parses (and writes) the OLE2-based binary formats used by Office
97 through 2003: Word (`.doc`), Excel BIFF8 (`.xls`), and PowerPoint (`.ppt`).
It builds on `litchi-cfb` for the CFB storage substrate and provides the
shared infrastructure those formats need: PLCF tables, SPRMs, and the
OfficeArt (Escher) drawing layer.

## Usage

```toml
[dependencies]
litchi-ole = "0.0.1"
```

```rust
use litchi_ole::XlsWorkbook;
use std::fs::File;

let file = File::open("example.xls")?;
let workbook = XlsWorkbook::new(file)?;
let sheet = workbook.xls_worksheet(0)?;
# Ok::<(), litchi_ole::XlsError>(())
```

## Features

- `.doc` (Word 97-2003) reader and writer with full PLCF/SPRM handling
- `.xls` (BIFF8) reader and writer with formula and shared-string support
- `.ppt` (PowerPoint 97-2003) reader and writer with Escher drawings
- Optional `formula` feature for MathType (MTEF) extraction
- Optional `imgconv` feature for EMF/WMF/PICT image bridges

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
