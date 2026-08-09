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

## Safety and mutation boundary

`Worksheet` is an authoring model. A worksheet decoded from an existing
workbook is source-bound because `litchi-xls` does not provide a whole-sheet
BIFF8 re-save path; its public create-only mutators return a typed refusal
instead of accepting an edit that could be dropped. Use an existing
feature-specific transaction where one applies, or create a new worksheet with
`Writer`.

`Writer::set_password` permits the RC4 profiles. Legacy BIFF8 XOR obfuscation
is intentionally decode-only by default and can only be authored with the
explicit `WeakEncryptionPolicy::allow_xor_obfuscation()` capability and
`Writer::set_xor_obfuscation_password`; it is not confidentiality protection.

```rust
use litchi_xls::Workbook;
use std::fs::File;

let workbook = Workbook::new(File::open("example.xls")?)?;
let _sheet = workbook.xls_worksheet(0)?;
println!("sheets: {}", workbook.sheets().len());
# Ok::<(), Box<dyn std::error::Error>>(())
```

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
