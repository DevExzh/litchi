# XLDM outer storage and XPRESS validation

This batch adds a fixture-backed Tabular150 outer-storage profile alongside canonical XLDM140, preserves borrowed source bytes for exact writes, validates allocation ranges and checksums, and separates bounded XPRESS decoding into explicitly selected WUSP32 and Tabular16 framing. It does not provide model evaluation or semantic model authoring.

Rust 1.95.0 validation passed 1,268 unit/integration tests and two doctests, strict all-target/all-feature Clippy, changed-file formatting, and whitespace checks. The native fixture regression checks all 48 allocation CRCs, exact borrowed source/write behavior, and bounded decoding of 46 logged members. An independent Python extractor verifies the pinned archive and model hashes, allocation CRCs, backup-log lengths, and 78 frame headers (75 compressed, three raw). The extractor does not decode XPRESS itself.

The fixture provenance and licenses are recorded beside the fixture. This is evidence for one pinned producer file and synthetic boundary tests; it supplies no timing, total-memory, or universal version-150 compatibility claim. `receipt.json` binds the exact source files and compressed logs.

XML marker decoding checks encoded length and UTF-16 expansion before reserving decoded bytes and does not build a temporary word vector. Boundary tests cover exact decoded size, expanded output refusal, and a valid-CRC oversized UTF-16 marker that must fail with the XML byte-limit error.

The upstream LGPL license is retained verbatim. Its two original trailing spaces are excluded from the whitespace-only check.
