# 0362ul120 — OOXML common VBA family error export

Status: source-only glossary repair from `a8d32513b8e427c6d234ddd17fb836b0e8d2477b`
(`uo2/common-glossary-0362ul120`).  This packet makes no compiler, Cargo,
rustdoc, runtime, native, or performance claim.

## Exact failure

The frozen common full-profile packet is retained at
`(deleted-coord-artifact)`.
Its source is the exact `a8d32513b8` tree.  The full rustdoc command exited
zero; its retained log is 19,696 bytes with SHA-256
`4e3b939741451d6e374b04fca54a70e800d79fe58ef3fc1570b0f17a300f7914`.
The subsequent glossary check exited one with the sole diagnostic below; its
73-byte log has SHA-256
`8ba0aa8b1815cad108d169be8b120bd0783ca952791cbca6f9cda228498bc87f`:

```text
ooxml-common-full: litchi_ooxml_common::vba requires exported type Error
```

The packet's source and result anchors remain unchanged: `rows.json` is
15,764 bytes with SHA-256
`e18dd131041cf56d2fe5c9fffb0ac826c702efeaf518f23515e1e792cd6d10de`,
and `result.json` is 17,095 bytes with SHA-256
`6c5886a33159a2f50d75f2ebfda9cd036e9a51f99aa4410b300717d3e85701c8`.

## Contract audit and repair

Record 0351 §2.3 requires every registered edit family to expose the
canonical six-name set:

```text
Edit, Patch, Commit, Diagnostics, Error, LimitKind
```

The `ooxml-common-full` profile registers `vba`, `custom_xml`, `properties`,
`web`, and `ribbon`.  The retained rustdoc graph shows that `custom_xml`,
`properties`, `web`, and `ribbon` already export all six names.  `vba` exports
`Edit`, `Patch`, `Commit`, `Diagnostics`, and `LimitKind`; its `Error` was only
the private item brought into scope by `use crate::{Error, Result}`.  The
glossary checker correctly rejects that private import rather than inferring a
module API from return types.

`litchi_ooxml_common::Error` is already the canonical shared OOXML error enum,
including the typed refusal, patch-conflict, signed, source-changed, internal,
and limit variants required by record 0351 and 0362ul94a.  The repair changes
the VBA prelude to:

```rust
pub use crate::Error;
use crate::Result;
```

This is the same nominal error used by every existing VBA operation, not a
new wrapper, spelling, alias, or compatibility path.  The DOCX, PPTX, XLSX,
XLSB, and facade consumers were audited: they import VBA operations and the
host/option types, while their own errors classify the common error at their
format boundary; none depended on a private `vba::Error` path.  Therefore no
consumer signature or behavior changes are required beyond exposing the
canonical family export.

No glossary exemption or weakened family check is added.  The common profile
will be rechecked by the coordinator against a fresh rustdoc graph after this
source packet is composed.

## Source-only checks

The retained rustdoc JSON was inspected for all five registered common
families and the exact `vba` child exports.  `git diff --check` and the staged
repository-size/binary guards are required before commit.  Cargo, compiler,
rustdoc, metadata, runtime, native, and performance execution remain deferred
to the coordinator.
