# litchi-xls

Typed reader and writer for the legacy Microsoft Excel BIFF (`.xls`) format.

## Ownership

`litchi-xls` is the concrete XLS format crate. It owns workbook, worksheet,
BIFF writer, chart-host, embedded-object, encryption, signature, and VBA
integration over the focused CFB, code-page, OfficeArt, OGraph, and common OLE
foundations. The former `litchi-ole::xls` path was deleted rather than retained
as a compatibility alias.

## Usage

```toml
[dependencies]
litchi-xls = "0.0.1"
```

```rust
use litchi_xls::XlsWorkbook;
use std::fs::File;

let workbook = XlsWorkbook::new(File::open("example.xls")?)?;
let _sheet = workbook.xls_worksheet(0)?;
println!("sheets: {}", workbook.sheets().len());
# Ok::<(), Box<dyn std::error::Error>>(())
```

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
