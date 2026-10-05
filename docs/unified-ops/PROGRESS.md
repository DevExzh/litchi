# Unified operations API — progress log

Driven by `docs/UNIFIED_OPS_GOAL.md`; specification `docs/UNIFIED_OPS_API_DESIGN.md`.

## Records

| No. | Title | Status |
|-----|-------|--------|
| 0001 | Owner override: read + write for every claimed format; tiers OOXML → OLE → (RTF, ODF, iWork) | recorded in the goal (2026-10-04) |
| 0002 | Leaf write-capability inventory (deep: OOXML + OLE; brief: tier 3) | merged (`cfe7681627`) |
| 0003 | M0a facade migration plan (flat `File`, `litchi::open`) | merged (`45a932ea2e`) |
| 0004 | ADR 0033 draft + `litchi-ops` scaffold + gate registration | merged (`af70966a6d`) |
| 0005 | M0b path/selector grammar in `litchi-core::selector` | merged (`9f8e49b181`) |
| 0006 | Owner accepts ADR 0033 | done |
| 0007 | Coordinator defaults for 0003's open questions | adopted, awaiting owner |
| M0a | Steps S1–S10 on `uo/m0a` (plan 0003 §4) | in progress |
| 0008 | Diagnosis: source-backed DOCX edits refuse every Word-authored fixture (MCE) — M2 blocker | merged (`e9c3868109`); fix staged as 0011 |
| 0009 | Design: editing a shared snapshot (R4; leaf publishers consume their package) — M2 prerequisite | in progress |
| 0010 | M1 unified-read plan (node view, path resolution, query, view modes, capabilities) | in progress |
| 0011 | DOCX MCE admission, stages 1–3 of 0008 (lands on `uo/m0a`) | in progress |

## Awaiting owner confirmation

- Record 0007 defaults: flat `File` (F2), `limit_exceeded`/`cancelled` codes (Q1), iWork → `unsupported_format` until tier 3 (Q2), `retain_source = true` (Q3), `Signed` wraps every signed input (Q4).

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

## 0003 — log

- Plan for M0a: a flat `File` enum with one variant per format, plus `Locked`/`Signed`, each variant one `Arc` handle (16 bytes, Send + Sync). The editable owner is built at the first edit, not at open. Leaf constructors L-PPTX-1, L-XLSX-1, L-XLSB-1, L-DOC-1, L-XLS-1 and L-PPT-1 avoid indexing the container twice.
- Gate: `non_iwork_gate.py` forbids any edge from a bulk package to an iWork crate. So `litchi-ops` stays free of iWork until M6, and the gate must be redesigned first (R1).
- Findings against the design: F1 (§16.2's "mutable owners" for pptx/xlsx are read-only), F2 (flat enum), F14 (Python `FileFormat` drift, including Odb → Doc).

## 0005 — log

- `litchi-core::selector` grammar: paths (1-based → `Position`, `last()`, quoted segments), predicates with `and`/`or`/`not`, header-row predicates, CSS-like queries (`:contains`, `:has`, `:not`, combinators, unions, sheet scopes, ranges), story axes, typed `ParseError` with offset and code, explicit limits.
- 950 OfficeCLI paths and 132 selectors parse and round-trip. Typical paths and simple queries parse with zero allocations (~57 ns per path, ~120 ns per selector, release build, single machine). There are 30 recorded OfficeCLI deltas: element aliases and bare-key meaning move to M1 resolution, and the unscoped-mutation guard moves to the verb layer.
- Open: the `selector_grammar` fuzz target compiles but has not been run (no nightly toolchain or cargo-fuzz on the host).

## 0008 — log

- None of the 19 DOCX fixtures uses a content-changing MCE construct. 13 only declare the namespace (Word 2007 `xmlns:ve`) and 6 carry one root `mc:Ignorable`.
- The source-backed guard refuses them because it compares buffers, and the MCE codec rewrites any part that mentions the namespace URI. The guard also has a hole: a namespace URI written with a character escape slips past both the codec and the guard.
- The eager owner keeps every `mc` construct byte for byte, but it has no MCE or protection check. It accepts a wrong-target edit under body-level `AlternateContent`, accepts `MustUnderstand` documents, and publishes edits to 3 protected fixtures.
- `to_owned_package()` rebuilds the ZIP layout (M0a `SourceMode::Eager` uses it). The execution-context route refuses even reads of all 19 files.
- Reads and edits count paragraphs differently: reads include table cells, edits count only top-level body paragraphs. M1 has to choose for `/body/p[N]`.
- Recommended: a staged admission rule (declared-only, 13/19; then Ignorable-only, 19/19) applied to both owners and the execution-context route, with one protection policy. Run-level `AlternateContent` comes later with a post-commit check. No silent eager fallback.

## Awaiting owner decision

- ~~ADR 0033 acceptance~~ — accepted 2026-10-04 (record 0006).
