# Unified operations API — progress log

Driven by `docs/UNIFIED_OPS_GOAL.md`; specification `docs/UNIFIED_OPS_API_DESIGN.md`.

## Records

| No. | Title | Status |
|-----|-------|--------|
| 0001 | Owner override: read + write for every claimed format; tiers OOXML → OLE → (RTF, ODF, iWork) | recorded in the goal (2026-10-04) |
| 0002 | Leaf write-capability inventory (deep: OOXML + OLE; brief: tier 3) | merged (`cfe7681627`) |
| 0003 | M0a facade migration plan (flat `File`, `litchi::open`) | in progress |
| 0004 | ADR 0033 draft + `litchi-ops` scaffold + gate registration | merged (`af70966a6d`) |
| 0005 | M0b path/selector grammar in `litchi-core::selector` | in progress |
| 0006 | Owner accepts ADR 0033 | done |

## Awaiting owner confirmation

- `OpenOptions` shape: one struct with common knobs plus optional per-format sub-options.
- Verb names: `apply`; find/replace as `find=`/`replace=` props on `set`; `merge` = template merge.

## 0004 — log

- Drafted ADR 0033 (Proposed) and listed it in the ADR index; ADR 0024 inventory amended (66 packages / 245 declarations).
- Scaffolded inert `litchi-ops` (edge to `litchi-core` only); registered in `crate_boundaries.json` and `non_iwork_gate.py` (bulk 46 → 47); gate tests updated and green.
- Known: workspace `cargo fmt --all --check` fails on base-branch files (soapberry-zip, iwa-archive) unrelated to this program.
- Follow-up: the first iWork leaf edge from `litchi-ops` (tier 3) requires reclassifying it out of the bulk gate, like the facade.

## 0002 — log

- Code-verified inventory; facade is read-only for every format and drops/never holds the editable leaf owner for doc, ppt, rtf; xlsx/xlsb are held only in the non-editable lazy form.
- No `Locked`/`Signed` types exist. Changed commits silently strip signatures in pptx `opened::Transaction`, several xlsb families, odf-formula, ODS legacy mutators, ODP in-place mutators (ADR 0006 concern; fix in the signature-policy prerequisite).
- doc/ppt `Writer::set_property` never emitted; xls `Writer` writes no property streams. Feature-matrix overclaims: pptx hyperlink write, docx app.xml write.
- Stream order: T1 xlsx > docx > pptx > xlsb; T2 xls > ppt > doc; T3 per record. Cross-cutting prerequisites: facade atomic writer, signature/encryption policy, commit-shape adapter, editable owner retention in M0a.

## Awaiting owner decision

- ~~ADR 0033 acceptance~~ — accepted 2026-10-04 (record 0006).
