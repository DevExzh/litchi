# 0362y — Keep the XLS eager-cell walk concrete

The legacy XLS adapter no longer stores its eager worksheet cell walk as a
private `Box<dyn Iterator>`. `litchi_xls::worksheet::StoredCells<'a>` names
the existing zero-copy `BTreeMap` values iterator, and
`format::xls::{owner,children}` carry that concrete iterator through the
cursor state. The worksheet's values, order, and lifetime are unchanged; no
cell data is collected or copied.

This closes the memo 29 residual without changing the source-backed cursor,
grid fallback, path behavior, or error classification. The source-backed XLS
commit path now retains its validated overlay plan and defers opening the
composed workbook until the first semantic read. Write-only publication still
streams the checked plan directly, while the first read opens the target once
and caches it for subsequent projections. Commit-time source, span,
fingerprint, and numeric readback validation remains unchanged. The
coordinator should qualify write-only, first-read, and read-after-write costs
against the former forward-only plan before claiming the F3 saving.

Validation was limited to formatting and static diff/artifact guards in this
leaf. The coordinator should run the XLS and unified-ops checks on the merged
assembly.
