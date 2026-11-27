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
`Builder`.

`Builder::set_password` permits the RC4 profiles. Legacy BIFF8 XOR obfuscation
is intentionally decode-only by default and can only be authored with the
explicit `WeakEncryptionPolicy::allow_xor_obfuscation()` capability and
`Builder::set_xor_obfuscation_password`; it is not confidentiality protection.

```rust
use litchi_xls::Package;

let package = Package::open("example.xls")?;
let workbook = package.workbook();
let _sheet = workbook.worksheet(0_usize)?;
println!("sheets: {}", workbook.worksheet_count()?);
# Ok::<(), Box<dyn std::error::Error>>(())
```

`Package` retains an immutable source and returns cheap owned read projections.
It publishes exact bytes through `to_bytes`, `write_to`, and atomic `save`.
Encrypted inputs require explicit source export; ordinary publication refuses
them. The old generic eager reader is available only as `raw::Reader`.
Primary changes use `Package::edit()` and `Package::apply(&patch)`. Fresh authoring
uses `Builder::build()` to produce a validated Package; publication belongs to
that owner. Builders have no save, stream or byte-output methods. Encrypted
builds retain the encrypted image and publish it explicitly through
`Package::to_encrypted_bytes()`; ordinary publication refuses encrypted owners.
`Package::open_with_password(path, password)` opens a protected source with the
default limits. `Package::encryption()` reports its retained `EncryptionProfile`,
including the CryptoAPI RC4 key length, after successful password verification.
Plain owners report `None`; this metadata does not authorize a security downgrade.

RC4 authoring requires explicit entropy for each build:

```rust
# fn main() -> Result<(), Box<dyn std::error::Error>> {
use litchi_xls::{Builder, EncryptionEntropy, EncryptionProfile};
let mut builder = Builder::new();
builder.set_password("secret", EncryptionProfile::OfficeBinaryRc4)?;
builder.set_encryption_entropy(EncryptionEntropy::generate()?);
let encrypted = builder.build()?.to_encrypted_bytes()?;
# Ok(())
# }
```

`build()` never obtains randomness. Each RC4 build attempt consumes its entropy,
even if later validation fails. Supply fresh entropy for the next attempt.
`EncryptionEntropy::from_bytes` supports reproducible pipelines; never reuse a
salt with the same password for different plaintexts. XOR obfuscation and plain
builds do not require entropy.


## Eager reader limits

`OpenOptions::source_limits` selects checked physical input, per-stream and
aggregate-stream ceilings. Defaults are 512 MiB, 512 MiB and 2 GiB. The reader
checks input size before CFB parsing and every declared stream before reading
format payloads, including opaque and nested streams. An already parsed OLE
handoff checks its catalog before decoding; it cannot undo earlier container
allocations. Invalid profiles and excess input have typed errors.

```rust
use litchi_xls::{OpenOptions, Package, SourceLimits};

let limits = SourceLimits::builder()
    .max_input_bytes(64 * 1024 * 1024)?
    .max_entry_bytes(16 * 1024 * 1024)?
    .max_aggregate_bytes(64 * 1024 * 1024)?
    .build()?;
let package = Package::open_with_options(
    "example.xls",
    OpenOptions::new().source_limits(limits),
)?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

The positional reader continues to use `SourceBackedLimits`; its explicit
materialization adapter preserves that existing admission policy. Focused
edit snapshots retain their separate object/CFB profiles. These limits are
not a total-memory budget.

## Primary edits

The source remains immutable. Commit exposes `package()`, `patch()` and
`diagnostics()`, or `into_parts()` for the `(Package, Patch)` pair. The package
retains the effective source and edit limits through no-op commits, further
edits and inverse application. Signed or encrypted inputs are refused by this
edit path. Unsupported BIFF dependency closures remain typed refusals.

```rust
use litchi_xls::{Package, cell_values::Reference};

let source = Package::open("example.xls")?;
let mut edit = source.edit()?;
edit.set_numeric("Sheet1", Reference::new(0, 0)?, 42.0)?;
let commit = edit.commit()?;
commit.package().save("edited.xls")?;
let restored = commit.package().apply(&commit.patch().inverse())?;
assert_eq!(restored.to_bytes()?, source.to_bytes()?);
# Ok::<(), Box<dyn std::error::Error>>(())
```


Minimal fresh workbooks use `Package::new()?`, which creates one empty `Sheet1`.
The returned owner supports primary transactions and atomic publication.

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
