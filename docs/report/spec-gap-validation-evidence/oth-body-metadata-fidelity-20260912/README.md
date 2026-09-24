# OTH metadata projection validation

The frozen source adds bounded namespace-aware metadata projection, strict
date/dateTime validation, and direct tracked-change ownership checks in the ODF
body prelude. Foreign descendants remain opaque. Existing source-preserving
edits, no-op behavior, inverse patches, and atomic failures remain covered.

Temporary and retained allocations are charged before reservation. Namespace
storage shares unknown URI values and rolls back bindings on element exit;
projection borrows the input instead of copying the complete XML. Pre-counted
collections reserve their exact measured capacity, including one-item vectors.
Repeated text is measured before allocation and repeated table structures stay
compact. Hash collection budgets are logical estimates, not allocator or RSS
measurements. The existing edit-only `block_owner` scan is outside this change.

Independent semantic and resource reviews approved this scope. The final
resource review specifically confirmed removal of the one-item geometric
reservation that exceeded its accounting. Earlier intermediate source hashes
are superseded by the final hashes in [root-gates.json](root-gates.json).

Root reran the gates against source hashes verified unchanged throughout:

- All-target tests: 98 passed (22 unit, 23 metadata, 11 structures, 33 API,
  and 9 edit tests).
- All-target Clippy with `-D warnings`: passed, without lint suppression.
- Rustdoc with `RUSTDOCFLAGS=-D warnings`: passed.
- Rustfmt checks for all six source/test files: passed.

The JSON records exact commands, toolchain, source hashes, and decompressed log
hashes. The four gzip files retain their complete command output. These are
correctness and bounded-resource checks, not runtime performance measurements.
Rich metadata coverage uses synthetic fixtures; the native OTH fixture is
empty and does not establish native rich-metadata interoperability.
