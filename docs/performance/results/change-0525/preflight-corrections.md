# Source review corrections before candidate capture

The initial production draft and focused tests were applied after all baseline
captures except the deliberately deferred native A2 repeat. No candidate Rust
build or runtime measurement had been run when these issues were corrected.

- `Rect::end()` is a two-dimensional exclusive bound. Same-row omissions must
  compare their end column in their own row; membership uses `Rect::contains`.
  The direct merge regression requires `Some`, preventing full-parser fallback
  from concealing broken ordering.
- Omission membership now advances monotonically through sorted ranges rather
  than scanning every range for every cell. Parsed changed records move into
  the combined store; only omitted source records are cloned.
- Reduced XML, parsed scratch, and omission ranges live in an optional helper.
  They drop before complete-parser fallback. The outer constructor retains one
  full XML validation and the existing source/execution fences.
- Ineligible rewrites render from the already scanned layout. Optional omission
  allocation failure discards provenance and retains the ordinary full output.
- The eligible value-only renderer no longer duplicates unreachable row, column,
  default, or root-effect rendering branches. `Plan::cells` cannot request those
  effects; the ordinary renderer remains authoritative for other edit kinds.
- Four raw byte marker scans were removed. The complete value-only validator
  already rejects foreign/unknown elements and unsupported attributes, including
  actual MCE and x14ac markup. Searching scalar text and unused namespace
  declarations for those words is neither necessary nor semantically precise.

These are source-review findings, not measured speedup claims. The canonical
candidate source manifest and runtime gates remain the authority for acceptance.
