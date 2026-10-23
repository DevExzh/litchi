# 0362y — Keep the XLS eager-cell walk concrete

The legacy XLS adapter no longer stores its eager worksheet cell walk as a
private `Box<dyn Iterator>`. `litchi_xls::worksheet::StoredCells<'a>` names
the existing zero-copy `BTreeMap` values iterator, and
`format::xls::{owner,children}` carry that concrete iterator through the
cursor state. The worksheet's values, order, and lifetime are unchanged; no
cell data is collected or copied.

This closes the memo 29 residual without changing the source-backed cursor,
grid fallback, path behavior, or error classification. The source-backed XLS
commit reopen optimization (0359a-F3) remains a separate follow-up because it
needs matched write-only, first-read, and read-after-write evidence while
retaining commit-time source validation.

Validation was limited to formatting and static diff/artifact guards in this
leaf. The coordinator should run the XLS and unified-ops checks on the merged
assembly.
