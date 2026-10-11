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

## Eager reader limits

`OpenOptions::source_limits` selects checked physical input, per-stream and
aggregate-stream ceilings. Defaults are 512 MiB, 512 MiB and 2 GiB. The reader
checks input size before CFB parsing and every declared stream before reading
format payloads, including opaque and nested streams. An already parsed OLE
handoff checks its catalog before decoding; it cannot undo earlier container
allocations. Invalid profiles and excess input have typed errors.

```rust
use litchi_xls::{OpenOptions, SourceLimits, Workbook};
use std::fs::File;

let limits = SourceLimits::builder()
    .max_input_bytes(64 * 1024 * 1024)?
    .max_entry_bytes(16 * 1024 * 1024)?
    .max_aggregate_bytes(64 * 1024 * 1024)?
    .build()?;
let workbook = Workbook::new_with_options(
    File::open("example.xls")?,
    OpenOptions::new().source_limits(limits),
)?;
# Ok::<(), litchi_xls::Error>(())
```

The positional reader continues to use `SourceBackedLimits`; its explicit
materialization adapter preserves that existing admission policy. Focused
edit snapshots retain their separate object/CFB profiles. These limits are
not a total-memory budget.

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
