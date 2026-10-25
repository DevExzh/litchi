# 0362ul68a — XLDM raw-policy review return

Status: source review return for 0362ul68. The policy remains ready for
independent build review; Cargo, native qualification and rustdoc remain
deferred to the coordinator's build slot.

## Error-boundary repair

The XLSB Data Model proof adapter now carries outer XLDM `Limit` and
`InvalidLimit`, section-2.5 metadata `Limit` and `InvalidLimit`, and OLAP
proof `Limit` outcomes into the package error taxonomy. The transaction's
metadata and OLAP proof paths use the same typed mappings. Genuine malformed,
unproven, unsupported-profile and allocation proof outcomes keep their prior
invalid/refused classification.

The XLSX Data Model boundary now preserves section-2.5 and OLAP-proof limit
dimensions as typed Data Model `LimitKind` variants, including proof item,
string, source and work bounds, instead of reporting them as unsupported
time-grouping content. The facade comment now names all three raw entry
points (`inspect`, `inspect_shared`, and `write`). Focused adapter tests cover
raw, metadata and OLAP-proof limit dimensions plus malformed-proof fallback.

## Retained output and XML-text repair

Metadata XML general references now charge their decoded bytes against the
cumulative `max_metadata_text_bytes` policy, including repeated named and
numeric references.

Same-size and variable-size identity rewrites enforce the inspected storage's
retained output ceiling before any output allocation. Variable rewrites use
the minimum of the retained and explicit caller ceilings, including their
empty-copy path. Identity rename errors retain typed output-limit dimensions;
an exact no-op still borrows the original source without copying. Focused
tests cover the entity-reference budget, retained output refusal and no-op
borrowing behavior.

## Source-only checks

The follow-up keeps the original 0362ul68 repository size and no-binary
guards. Formatting and diff checks are run before commit. No Cargo command,
rustdoc command, native fixture run, benchmark, or performance qualification
is claimed here.
