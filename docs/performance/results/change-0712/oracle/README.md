# 0712 active-offset diagnostic probe

This standalone binary is a bounded diagnostic for the public
`litchi_docx::alt::{scan, active}` APIs and the public mutable DOCX facade. It
does not measure time, allocation, or throughput. Each synthetic document is
hashed and records its byte length, declared relationship IDs, every known
synthetic `w:p`, `w:tbl`, and `w:altChunk` source start, the complete `scan`
outcome, and three exact `active` outcomes:

* `active_full` receives every synthetic Word block start;
* `active_all_alt` receives only `altChunk` starts; and
* `active_empty` receives no offsets.

Each outcome preserves both `Display` and `Debug` error text. The offset input
count, little-endian offset SHA-256, and short input vector (when at most 64
items) are recorded as well. The `document_mut` field constructs a minimal
OPC package and invokes `Package::from_opc_package(...).document_mut()`, which
exercises the same public writer admission path that first scans anchors and
then selects active document ranges.

The cases include transitional and strict namespaces, ordinary no-anchor
paragraph/table blocks, valid MCE choice and fallback branches, inactive
anchors, malformed `AlternateContent` with no `Choice`, an unknown
`MustUnderstand` namespace, and anchors nested inside a paragraph and table.
The nested cases intentionally record the distinction between `active`, which
can select every supplied source start, and the writer's range scanner, which
captures an outer target range and therefore suppresses nested target ranges.

The crate-private implementation limit is currently 1,000,000 visibility
offsets. The final case supplies 1,000,001 repeated offsets to the public
`active` function without creating a million-element XML document; the report
records the resulting observed error rather than assuming its text.

Run from the repository root after seeding the lockfile from the parent
experiment:

```text
cargo run --release --locked --manifest-path \
  docs/performance/results/change-0712/oracle/Cargo.toml -- \
  --output docs/performance/results/change-0712/oracle/report.json
```

The output path is not included in the report. No result is interpreted as a
pass by the probe; acceptance analysis belongs to the parent experiment after
the captured JSON has been inspected.
