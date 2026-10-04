# Unified operations API — progress log

Driven by `docs/UNIFIED_OPS_GOAL.md`; specification `docs/UNIFIED_OPS_API_DESIGN.md`.

## Records

| No. | Title | Status |
|-----|-------|--------|
| 0001 | Owner override: read + write for every claimed format; tiers OOXML → OLE → (RTF, ODF, iWork) | recorded in the goal (2026-10-04) |
| 0002 | Leaf write-capability inventory (deep: OOXML + OLE; brief: tier 3) | in progress |
| 0003 | M0a facade migration plan (flat `File`, `litchi::open`) | in progress |
| 0004 | ADR 0033 draft + `litchi-ops` scaffold + gate registration | in progress |
| 0005 | M0b path/selector grammar in `litchi-core::selector` | in progress |

## Awaiting owner confirmation

- `OpenOptions` shape: one struct with common knobs plus optional per-format sub-options.
- Verb names: `apply`; find/replace as `find=`/`replace=` props on `set`; `merge` = template merge.
