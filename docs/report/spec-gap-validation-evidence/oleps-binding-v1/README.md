# OLE Property Set binding bit order

GUID-derived stream names now follow the packet-byte bit order in local MS-OLEPS §2.23. The previous encoder and decoder agreed with each other but generated names incompatible with the published PropertyBag example in §3.2. The corrected codec emits the canonical lowercase name, accepts the case-insensitive normative spelling, and rejects nonzero trailing bits. PROPERTY_BAG_FMTID is exported alongside the existing standard identifiers.

The regression binds the normative FMTID to its published name and checks a second fixed packet vector, uppercase decoding, malformed characters, and trailing bits. Independent review checked the GUID packet representation against local MS-DTYP. Root isolation passed 208 unit/integration tests, with one existing native-corpus test ignored, strict all-target/all-feature Clippy, formatting, and whitespace checks using Rust 1.95.0. There are no doctests in this crate. No performance or new storage-authoring claim is made.

This corrects generic binding names; previously emitted incorrect names are not introduced as alternate aliases. The receipt binds the five source files and compressed validation logs. Specs were read from `3rdparty/specs/`.
