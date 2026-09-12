# OPC admission and source relationship handles

OPC package admission now refuses ZIP encryption declarations in either central
or local headers, including directory entries and strong-encryption flag bit 6.
The eager, lazy, and positional package paths enforce this before semantic part
materialization. Physical target reads refuse an encrypted target while retaining
the existing behavior for unrelated encrypted entries; the underlying ZIP reader
does not gain an unconditional package-wide policy.

Positional admission reads a bounded eight-byte local-header prefix per physical
entry. Successful validation is cached under the existing `ReaderAt` and
`IndexedArchive` lifetime byte-stability contract. A counted-reader regression
checks one prefix request per entry during the first validation and no further
requests during the second. This is deterministic I/O-count evidence, not a
latency or throughput measurement. Source-backed owners retain their source and
execution fences around admission and subsequent operations.

Independent security review approved the final source revision, including local
and central encryption flags, directory entries, target-specific physical reads,
and the positional validation cache's existing byte-stability prerequisite.

`SourceBackedPackage::relationships_data_for_with_limit` retains exact canonical
relationship-member bytes in `PartData`. It clamps the caller cap to the package
cap before decompression, reserves managed memory/object resources, and charges
declared decompression work before reading. Handles retain reservations until the
last clone is dropped; work remains cumulative. Absent members are distinguished
from explicit empty members and return only after source/execution checks.
Read failures give source-change and cancellation errors precedence over archive
errors. Independent source/resource review approved this helper, including the
final formatting-only revision.

Root validation against unchanged hashes passed:

- Full default test suites and doctests for `litchi-opc` and `soapberry-zip`:
  1,102 passed, 3 ignored, none failed.
- All-target Clippy with `-D warnings`, without lint suppression.
- Rustdoc with `RUSTDOCFLAGS=-D warnings`.
- Rustfmt for all seven changed source and test files.

[root-gates.json](root-gates.json) records commands, toolchain, source hashes,
test summaries, and hashes of the decompressed logs. The gzip files retain full
command output. These gates do not approve the separately in-progress XLSX
form-control or pivot owner APIs, add ZIP decryption, or establish a performance
speedup.
