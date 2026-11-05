# 0351 — API contract: immutable packages, one edit/commit/refusal shape, configured limits

- Status: contract and foundation implemented on `uo2/api-contract`, cut
  from `feat/unified-ops-api` at `5ff605692e`. The format crates are not
  migrated here; the waves in §0 do that.
- Date: 2026-10-20.
- Input: owner decision [0350](0350-owner-api-debt-decision.md) (memo
  [0331](0331-owner-decision-memo.md) items 14–19, record
  [0343](0343-glossary-closure.md) §4 D1–D6), ADR 0034 §2–§6 and §9,
  `docs/API_GLOSSARY.md`.
- Owned: `litchi_core::edit` (new), the refusal classification in
  `crates/litchi-ops/src/error.rs`, `tools/check_api_glossary.py` and its
  tests, `tools/api_glossary.json`, `tools/api_debt_inventory.py` (new),
  `docs/API_GLOSSARY.md` (§2a, the D1–D6 cells, §5, §11), this record and
  `results/0351/`.
- Not owned (parallel records, treated as given): iWork `At`/`Area` and
  `litchi-iwa` builder names (0352), static dispatch of the
  `litchi-core::sheet` traits (0353), `litchi::sheet::text` move and
  `Metadata.keywords` (0354).

## 0. Wave table

Every wave depends on F0 (this record) being merged into trunk. At most
three waves run at once; the rounds below never put two waves on the same
crate. Sizes are counted from the public surface (§2) and source call sites;
XL is more than about 1,000 changed call sites, L 300–1,000, M under 300.

| Round | Wave | Package | Owned paths | Size | Depends on |
|---|---|---|---|---|---|
| — | F0 | Foundation (this record) | `crates/litchi-core/src/edit.rs`, refusal arm of `crates/litchi-ops/src/error.rs`, `tools/check_api_glossary.py`, `tools/api_glossary.json`, `tools/api_debt_inventory.py`, glossary §2a | S (landed) | — |
| 1 | W1 | DOCX | `crates/litchi-docx/**`; `crates/litchi-ops/src/format/docx.rs`, `format/docx/**`, the docx `Classify` impls in `crates/litchi-ops/src/error.rs`, `crates/litchi-ops/tests/docx*`; docx hunks of `crates/litchi/src/**`, `crates/litchi/examples/docx*`, `crates/pyo3-litchi/src/{document,write}.rs`; `tools/perf-baseline/src/{docx_managed_edit,security_corpus}.rs`, `tools/perf-baseline/src/bin/docx_*`; `crates/litchi-docx/fuzz`; `docx-*` profiles; glossary §2a line W1 | XL: 85 `Package` + 44 `MutableDocument` mutators, 202 draft mutators, 30 commit and 26 edit types, 16 `Transaction`, 50 `*_with_limits`, `UnsafeEdit` at 194 sites in 50 files | F0 |
| 1 | W2 | PPTX | `crates/litchi-pptx/**`; `crates/litchi-ops/src/format/pptx.rs`, `format/pptx/**`, pptx `Classify` impls, `crates/litchi-ops/tests/pptx*`; pptx hunks of facade, `crates/pyo3-litchi/src/presentation.rs`; `tools/perf-baseline/src/bin/pptx_*`; `pptx-*` profiles; line W2 | XL: 77 `Package` mutators, 39 draft mutators, 25 commit and 19 edit types, 37 `*_with_limits`, `UnsafeEdit` at 224 sites in 52 files | F0 |
| 1 | W3 | PPT | `crates/litchi-ppt/**` (not `litchi-ole-common`, `litchi-ograph`, `litchi-crypto`); `crates/litchi-ops/src/format/ppt.rs`, `format/ppt/**`, ppt `Classify` impls; `ppt-*` profiles; line W3 | L: 14 `Transaction` types, 15 `rollback`, 16 `is_changed`, 20 commit types, 97 `*_with_limits`, `raw::Reader`, `writer::Chart` | F0 |
| 2 | W4 | XLSX, XLSB, OOXML common | `crates/litchi-xlsx/**`, `crates/litchi-xlsb/**`, `crates/litchi-ooxml-common/**`; `format/{xlsx,xlsb}*`, `format/workbook/**` in litchi-ops, their `Classify` impls; facade and `crates/pyo3-litchi/src/sheet.rs` hunks; `tools/perf-baseline/src/bin/{xlsx_*,xlsb_*}.rs`; profiles; line W4 | XL: 35 XLSX `Package` + 50 XLSB `Workbook` mutators, `into_workbook`, 61 draft mutators, 53 commit and 59 edit types, 20 `Transaction`, 35 `is_changed`, 139 `*_with_limits` | W1, W2 (shared `litchi-ooxml-common`), 0353 merged |
| 2 | W5 | DOC, XLS, RTF and legacy shared | `crates/litchi-{doc,xls,rtf,ograph,ole-common,crypto}/**`; `format/{doc,xls,rtf}*`, `ole_metadata.rs` in litchi-ops, their `Classify` impls; `tools/perf-baseline/src/bin/xls_*`; profiles; line W5 | L: DOC `raw::Package` (24), RTF `raw::Document` (163), 14 DOC `Transaction`, 38 commit types, 56 `*_with_limits`, XLS `UnsafeEdit(String)` at 587 sites in 50 files, ograph `UnsafeEdit` and `rollback`, 10 `litchi-ole-common` and 1 `litchi-crypto` `rollback` | W3 (PPT re-exports these crates), 0354 merged |
| 2 | W6 | ODF | `crates/litchi-{odf-common,odt,ods,odp,odg,odc,odi,odm,oth,odb,odf-formula}/**`; `format/{odt,ods,odp,odf_*}*` in litchi-ops, their `Classify` impls; `crates/litchi/examples/odt_*`; profiles; line W6 | L: ODT `MutableDocument` (136) and `generic::Package`/`odc::Document` (23), 13 ODS `rollback`, 54 ODS and 66 ODI draft mutators, 40 commit and 39 edit types, flat owners, 42 `*_with_limits` | F0, 0353 and 0354 merged |
| 3 | W7 | iWork | `crates/litchi-{pages,keynote,numbers}/**`, `format/iwork/**` in litchi-ops; profiles; line W7 | M: `into_parts` on 93 commit types, `MergeReader`/`Workbook` `*_with_limits` (4) | F0, 0352 merged |
| 3 | W8 | Unified `File` and closure | `crates/litchi-ops/src/{transaction,batch,import,merge,refresh}.rs`, `detect/**`, `crates/litchi/src/**`, `crates/litchi-cli/**`, `crates/pyo3-litchi/**`, ops profiles, glossary §2a line W8 and the final sweep | M: 16 `*_with_limits`, `transaction::Commit`/`BatchCommit` public fields | W1–W7 |

The shared litchi-ops dispatch files (`read`, `write`, `transaction`, `save`,
`create`, `capabilities`, `file`) and `Cargo.toml` change only in small,
format-separated hunks in W1–W7. Each wave adds its own `impl Classify`
arms in `error.rs` and touches no other format's arms.

## 1. Method

All-features rustdoc JSON was built for the 25 format and shared crates
(`litchi-core`, `-ooxml-common`, `-docx`, `-pptx`, `-xlsx`, `-xlsb`, `-doc`,
`-xls`, `-ppt`, `-rtf`, the eleven ODF crates, `-numbers`, `-pages`,
`-keynote`, `-iwa`) and `litchi-ops`, and walked with the gate's own
`Surface` class (`tools/check_api_glossary.py`), so the inventory sees
exactly what the gate sees. Each public item is counted once, at its
shortest public path (0343 counted re-export paths, which is why its PPT
`rollback` and `*_with_limits` figures, 19 and 173, are higher than the
item counts here). The generator is `tools/api_debt_inventory.py` (its
docstring has the build commands); its output for trunk `5ff605692e` is
[`results/0351/inventory.tsv`](results/0351/inventory.tsv) (325,260 bytes,
SHA-256 `67204950bb0ba0a1429fe434867a965b8b07d75e2fd4b1614a52299db5de62bf`,
one row per item with a target column). Source call-site counts are `grep -w`
over `crates/` and `tools/`.

| crate | mut_owner | mut_draft | commit_type | commit_snapshot | edit_type | transaction_type | rollback | is_changed | unsafe_edit | with_limits | into_workbook |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| docx | 129 | 202 | 30 | 37 | 26 | 16 | 1 | 16 | 1 | 50 | 0 |
| pptx | 77 | 39 | 25 | 24 | 19 | 0 | 0 | 0 | 1 | 37 | 0 |
| xlsx | 35 | 0 | 36 | 34 | 56 | 20 | 0 | 33 | 0 | 72 | 0 |
| xlsb | 50 | 61 | 17 | 16 | 3 | 1 | 0 | 2 | 0 | 50 | 1 |
| ooxml-common | 0 | 0 | 2 | 2 | 2 | 1 | 0 | 1 | 0 | 17 | 0 |
| ppt | 3 | 16 | 20 | 19 | 14 | 14 | 15 | 16 | 0 | 97 | 0 |
| doc | 24 | 0 | 22 | 25 | 15 | 14 | 0 | 19 | 0 | 30 | 0 |
| xls | 2 | 12 | 15 | 17 | 9 | 1 | 0 | 0 | 1 | 10 | 0 |
| rtf | 163 | 0 | 1 | 0 | 1 | 0 | 0 | 0 | 0 | 16 | 0 |
| odt | 159 | 13 | 3 | 2 | 6 | 2 | 1 | 0 | 0 | 3 | 0 |
| ods | 0 | 54 | 26 | 24 | 23 | 9 | 11 | 2 | 0 | 12 | 0 |
| odp | 0 | 0 | 4 | 4 | 5 | 2 | 1 | 0 | 0 | 0 | 0 |
| odg | 0 | 1 | 2 | 2 | 2 | 0 | 0 | 0 | 0 | 0 | 0 |
| odc | 0 | 2 | 3 | 4 | 2 | 0 | 0 | 0 | 0 | 2 | 0 |
| odi | 0 | 66 | 2 | 2 | 1 | 1 | 0 | 0 | 0 | 0 | 0 |
| odm / oth / odb / odf-formula | 0 | 0 | 3 / 1 / 1 / 1 | 0 | 0 / 0 / 1 / 0 | 0 | 0 | 0 | 0 | 0 / 0 / 0 / 1 | 0 |
| odf-common | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 24 | 0 |
| keynote / numbers / pages | 0 | 0 | 42 / 27 / 24 | 0 | 0 | 0 | 0 | 0 | 0 | 0 / 1 / 3 | 0 |
| ops | 0 | 0 | 2 | 0 | 0 | 0 | 0 | 0 | 0 | 16 | 0 |
| **total** | **642** | **466** | **309** | **212** | **187** | **81** | **29** | **89** | **3** | **444** | **1** |

Categories: `mut_owner` is a public `&mut self` method on a `Package`-like
owner; `mut_draft` one on a `Mutable*`, `Flat*`, `Reader` or projection-named
type; `commit_type` a commit missing a triad accessor or still publishing a
snapshot; `edit_type` a staging type named `*Transaction`/`*Edit` (other than
`Edit`) or exposing `rollback`/`is_changed`/`snapshot`. Two heuristic rows
are not debt: `litchi_core::SubEdit` (a durable-patch sub-edit, not a staging
type) and `litchi_iwa::numbers::UnverifiedTopologyEdit` (0352's crate). Not
counted by rustdoc: `UnsafeEdit` call sites — DOCX 194 in 50 files, PPTX 224
in 52, XLS 587 in 50, `litchi-ograph` 4 in 3, `litchi-ops` 4 in 2,
`tools/perf-baseline` 4 in 4.

Findings beyond the 0343 description:

- `rollback` also survives outside PPT: 11 ODS families (`&mut self`
  resets), ODP `charts::Edit`, ODT `header_footer::Transaction`, DOCX
  `chart::Transaction`, ten `litchi-ole-common` and one `litchi-crypto`
  transaction, and `litchi-ograph`. PPT's `rollback(self) -> Snapshot` only
  returns the owned base: outside tests, the only callers are two PPT
  delegations (`master::metadata::{Editor, template::Editor}`), so no writer
  depends on it (§3.5).
- `UnsafeEdit` also exists in XLS (`UnsafeEdit(String)`, a formatted string
  at 587 sites, which can carry authored content) and `litchi-ograph`.
- `*_with_limits` reaches the unified layer itself: `File::{edit, batch,
  import, merge, refresh}_with_limits` and nine `detect::*_with_limits`.
- XLSX scoped entries take `&mut self` (`edit_page_breaks`,
  `edit_custom_data`, …), not only the `put_*` mutators.
- PPTX keeps alias pairs `put_read_only_recommended`/`put_readonly_recommended`
  (and the matching `remove_*`/`apply_*`), which ADR 0034 §1 forbids.
- The `get_*` getters of DOC/XLS/PPT `OpenOptions` differ from the
  `*_value()` getters of RTF, ODF and iWork.
- ODT `mutable::MutableDocument` is a second public write path for opened
  documents (`MutableDocument::from_document`), not only the authoring model
  behind `Builder`.

## 2. Normative contract

"Must" rules are gate-checked where §4 says so and test-checked otherwise.
Every rule applies to every format crate and to the unified `File`.

### 2.1 Immutable owner (item 14)

1. A format's `Package` and its read-only projections (`Document`,
   `Workbook`, `Presentation`, `Drawing`, `Chart`, `Image`, `Database`,
   `Formula`, and the flat owners) expose no public method taking
   `&mut self`, inherent or through a trait impl. Gate: `immutable_owners`.
2. Outputs take `&self`: `save`, `save_with_durability`, `save_plain*`,
   `save_encrypted*`, `save_reencrypted*`, `to_bytes`, `to_plain_bytes`,
   `to_encrypted*`, `to_reencrypted*`, `write_to`. An output that today
   flushes a hydrated model or caches serialized bytes keeps the cache behind
   interior mutability or serializes into a local buffer; repeated output of
   one package is byte-identical.
3. Signature verbs return a new package: `sign(&self, ..)`,
   `resign(&self, ..)`, `unsign(&self) -> Result<Package>`; the variants with
   options keep the `_with` suffix only where 0034 §4 already names them.
4. Every other mutator moves into a scoped `Edit` family (§2.3).
   `custom_props_mut` is deleted (record 0342's
   `metadata::Edit::set_custom_props` replaces it); `edit_opc(&self)` returns
   the raw OPC `Edit` whose commit publishes a `Package`; history verbs
   (`undo_document`, `redo_document`, `publish_document_commit_with_history`)
   become `&self -> Result<Commit>`.
5. XLSB: `Package::into_workbook` is deleted. `xlsb::Workbook` becomes the
   read-only projection (`Package::workbook()`); its 50 mutating verbs move to
   `Package::edit_x()` families, the same split XLSX made in 0250.
6. A `raw` module may keep a mutable codec model, but not under an owner
   name: DOC `raw::Package` (lazy `&mut self` readers) takes `&self` with
   interior caching or becomes `raw::Reader`; RTF `raw::Document` becomes
   `raw::Tree` (or another non-owner noun) if it stays mutable.
7. Fresh-authoring drafts are not owners, but they follow the focused-module
   rule: no `Mutable*` names. Drafts reached from a `Builder` live in its
   module without the prefix (`docx::builder::{Document, Paragraph, Run,
   Table, Row, Section, Style}`, `pptx::builder::Slide`,
   `xlsb::builder::Worksheet`, `ods::builder::Workbook`); writer-side chart
   drafts (`ppt::writer::Chart`, `xls::chart::Chart` setters) move into their
   builder modules the same way. Gate: the `Mutable` prefix ban.
8. Cost: `Package::edit*()` is O(1) in the package size (the package is a
   cheap shared snapshot); a commit shares every untouched part with its
   source.

### 2.2 The shared shapes (`litchi_core::edit`, landed in F0)

```rust
pub struct Diagnostics { /* changed: bool */ }          // Copy
impl Diagnostics { UNCHANGED, CHANGED, new(bool), changed(&self) -> bool }

pub struct Commit<P, T, D = Diagnostics> { /* package, patch, diagnostics */ }
impl Commit { new(P, T, D), package(&self) -> &P, patch(&self) -> &T,
              diagnostics(&self) -> &D, into_package(self) -> P,
              into_parts(self) -> (P, T) }

#[non_exhaustive] pub enum RefusalKind { UnsupportedOperation,
    UnsupportedContent, UnsupportedTarget, UnsupportedProperty,
    InvalidValue, NotFound, Protected }                  // ALL, as_str
pub struct Refusal { /* kind, operation: &'static str, reason: &'static str */ }
impl Refusal { new(kind, operation, reason), kind(), operation(), reason() }
```

The shared types live in `litchi-core` because every format crate already
depends on it (`litchi-ooxml-common` would reach only the four OOXML crates).
A family whose commit is exactly the triad writes
`pub type Commit = litchi_core::edit::Commit<Package, Patch>;` and, when it
reports only the change flag, `pub use litchi_core::edit::Diagnostics;`. A
family with more (XLSX per-edit changes, iWork touched components) keeps its
own types with the same accessors. Both forms pass the gate.

### 2.3 Edit, commit, patch (items 15, 16)

The ADR 0034 §3 shape, now without exceptions:

| Step | Primary transaction | Scoped family `x` |
|---|---|---|
| Read | `Package::document()` / `workbook()` / `presentation()` | `Package::x()` |
| Stage | `Package::edit(&self) -> Edit` | `Package::edit_x(&self) -> x::Edit` |
| Configured stage | `Package::edit_with_options(&self, EditOptions)` | `Package::edit_x_with_options(&self, x::EditOptions)` |
| Publish | `Edit::commit(self) -> Result<Commit>` | `x::Edit::commit(self) -> Result<x::Commit>` |
| Replay | `Package::apply(&self, &Patch) -> Result<Commit>` | `Package::apply_x(&self, &x::Patch) -> Result<x::Commit>` |
| Result | `Commit::{package, into_package, patch, diagnostics, into_parts}` | same on `x::Commit` |
| No-op | `commit.diagnostics().changed() == false`, `patch.is_noop()`; the package is the source unchanged | same |
| Undo | `patch.inverse()` applied with `apply`/`apply_x` | same |
| Discard | drop the `Edit`; `Edit::reset(&mut self)` clears staged changes in place | same |

Rules:

1. Types live in a focused module, `x::{Edit, Patch, Commit, Diagnostics,
   Error, LimitKind}`; no prefixed names (`FlatDrawingEdit`,
   `FontTransaction`, `PictureBulletTransaction`, `SourceTransaction`,
   `WorksheetEdit`, …). Gate: `families` + `family_methods`; the
   `Transaction` exact and suffix bans.
2. No `Transaction`, `rollback`, `is_changed`, or `Commit::{snapshot,
   into_snapshot, workbook, into_workbook}`. A commit always publishes the
   format's `Package`, never a part-level snapshot.
3. Part-level snapshots may remain as read views (`Package::custom_xml() ->
   custom_xml::Snapshot`), but the publication path that took one back
   (`put_x(snapshot)`, `apply_x_commit(commit)`, `store_*`, `take_*`) is
   deleted: the commit already carries the package. `apply_x_plan(plan)`
   becomes an `Edit` verb staging the plan.
4. The ~200 scoped OOXML families (DOCX 30 commit / 26 edit types, PPTX
   25/19, XLSX 36/56, XLSB 17/3, OOXML common 2/2, inventory rows
   `commit_type` and `edit_type`) each move their staging onto the package:
   the `Edit` holds a cheap clone of (or reference to) the source package and
   its staged part changes; `commit` validates, applies them to the OPC
   overlay and returns the new `Package` with its `Patch` and `Diagnostics`.
   Their old `put_*`/`remove_*`/`apply_*_patch` package mutators are the
   same families' entry points (tables in §3).
5. Mutation of a signed package refuses a changed commit with the module's
   `Error::Signed`; an exact no-op commit preserves the signature
   byte-identically (ADR 0034 §9, unchanged).

### 2.4 Error taxonomy (item 17)

Every module error that a staging or apply path returns uses these
variants, with these payloads and litchi-ops codes:

| Variant | Meaning | litchi-ops code |
|---|---|---|
| `Refused(litchi_core::edit::Refusal)` | well-formed request this route refuses; classified by `RefusalKind` (§2.2 table) | `unsupported_capability`, `unsupported_type`, `unsupported_property`, `invalid_value`, `not_found`, `document_protected` |
| `PatchConflict { .. }` (unit allowed) | the patch or staged plan was made against another base | `invalid_input` |
| `SourceChanged { expected, observed }` | the source file changed after open | `io_error` |
| `Limit { kind, observed, maximum }` | a configured budget was exceeded | `limit_exceeded` |
| `InvalidLimit { kind, value, maximum }` | an invalid configured limit | `invalid_argument` |
| `Signed` | a changed write on a signed package | `document_protected` |
| the module's internal variant | a broken invariant | `internal_error` |

1. `UnsafeEdit` is deleted everywhere (DOCX, PPTX root and `opened`, XLS,
   `litchi-ograph`). Each site maps by what its reason says:
   - "changed after/during planning or staging", "does not match the patch
     source", "unexpected serialized package revision", stale graph →
     `PatchConflict`;
   - markup-compatibility branches, unknown or nested children, opaque
     records, a field inside replaced text → `Refused(UnsupportedContent)`;
   - "supports one atomic operation", "managed transactions only support …",
     "no retained source-preserving implementation", "requires
     package-level publication" → `Refused(UnsupportedOperation)`;
   - "has no text body / notes page / shape properties", "takes part in a
     merge", "already has a title placeholder" → `Refused(UnsupportedTarget)`;
   - "document or write protection is enforced" → `Refused(Protected)`;
   - a missing admission, ledger, identity or authorization the code itself
     should have provided → the internal variant, not a refusal.
2. Typed reason enums that already exist (`docx::document::Refusal`,
   `doc::body_text::Refusal`, `ppt::{slide_order, text_edit}::Refusal`,
   `ooxml_common::mce::admission::Refusal`, the source-backed DOCX ones) stay
   as the payload of their module's `Refused` variant and gain
   `fn kind(&self) -> litchi_core::edit::RefusalKind`, so litchi-ops still
   classifies by kind alone. A crate with only free-text reasons uses the
   shared `Refusal`.
3. A refusal is content-free by construction: `&'static str` operation and
   reason. `UnsafeEdit(String)` sites in XLS move to static reasons; the
   redaction rule of ADR 0034 §6 applies to every message changed here,
   although the comprehensive audit (item 22) stays deferred.
4. `PatchConflict` is `invalid_input` in every format. Today DOCX
   `document::Error`, DOCX `metadata::Error` and the PPTX root map it to
   `io_error` (0343 kept that); W1/W2 move them, because with immutable
   packages a patch conflict is a caller supplying a patch for another base,
   while a changed file is `SourceChanged`.
5. ODT `MutableDocument`: the `mutable` module becomes crate-private. Its
   opened-document verbs (the 136 in the inventory) are reached only through
   `odt::Package::edit()` (the `Edit` already drives `MutableDocument`
   internally) and its fresh-authoring verbs through `odt::Builder`;
   `Content`/`ContentMut`/`Styles`/`StylesMut` follow. ODT
   `generic::Package` (12 `&mut self`) and `odc::Document` (11) move into
   `Edit` families like any other owner.

`litchi-core` holds the refusal vocabulary, not a shared error enum: each
module keeps its own non-exhaustive `Error` (ADR 0034 §6). The litchi-ops
`Classify` impl for `Refusal`/`RefusalKind` landed in F0; each wave adds
`Self::Refused(refusal) => refusal.classify()` arms for its crates.

### 2.5 Configured limits (items 18, 19)

1. One `OpenOptions` per format carries every ingress setting
   (`source_limits`, `semantic_limits`, leniency, encryption and cache
   limits, execution context). The configured entries are
   `open_with_options(path, OpenOptions)`,
   `from_bytes_with_options(&[u8], OpenOptions)`,
   `from_reader_with_options(R, OpenOptions)` and, where zero-copy sharing
   exists, `from_shared_bytes_with_options`. The ODF family's envelope is
   `litchi_odf_common::OpenOptions`, already re-exported by every ODF crate.
2. A package keeps the limits it was opened (or built) with. Edits and
   outputs use them; per-call configuration is `edit_with_options` /
   `edit_x_with_options` with an `EditOptions` whose defaults are the
   package's limits (the DOC and XLSB precedents, 0192 and 0261). Output
   verbs take no limits (`save_encrypted_with_limits` →
   `save_encrypted`, with the encryption limits from `OpenOptions` or the
   `Builder`).
3. No public `*_with_limits` name. An operation has one public entry: on an
   owner it takes no limits argument (rules 1–2); on a raw or record API
   (`ppt::raw::Reader`, record parsers, `detect::*`, codec functions) it
   takes `&SourceLimits` (or the module's limits) as a required argument,
   with no limit-less twin. Record parsers that no consumer outside the
   crate uses become `pub(crate)`. Internal functions keep taking limits.
   Gate: the `_with_limits` suffix ban.
4. PPT `raw::Reader`: `Reader::open_with_options(path, OpenOptions)` and
   `from_reader_with_options`, plus the existing default constructors; its
   `presentation_with_limits`-style methods go away, the reader's limits
   apply.
5. Flat ODF owners live at `flat::<Projection>` (the ODT precedent
   `odt::flat::Document`): `ods::flat::Workbook`, `odp::flat::Presentation`,
   `odg::flat::Drawing`, `odc::flat::Chart`, `odi::flat::Image`, each with
   `open_with_options`, `from_bytes_with_options` and
   `from_reader_with_options` taking `litchi_odf_common::OpenOptions`.
   `FlatDrawing` and `FlatPresentation` lose their fixed ceilings. The flat
   crate `Limits` types (`odt::flat::Limits`, `ods::flat::Limits`) and the
   ODC chart `Limits` fold into `SourceLimits`/`SemanticLimits` dimensions
   (new dimension nouns registered in the glossary; an existing dimension
   keeps its one name). Their prefixed staging types (`FlatDrawingEdit`,
   `FlatDrawingCommit`, `FlatDrawingPatch`, ODI `FlatImageTransaction`)
   become `flat::{Edit, Commit, Patch, ..}`.
6. Options accessors: a consuming setter named by the noun
   (`source_limits(self, SourceLimits) -> Self`) and a getter
   `<noun>_value(&self)` (the RTF, ODF and iWork form). DOC/XLS/PPT
   `get_source_limits`, `get_leniency`, `get_password`,
   `get_compatibility_profile` and DOC `EditOptions::with_*` setters are
   renamed.
7. Unified `File` (W8): `File::edit_with_options(EditOptions, f)` replaces
   `edit_with_limits`, and `batch`, `import`, `merge` and `refresh` take the
   same `litchi_ops::EditOptions` (`semantic_limits`) through their
   `*_with_options` forms; the `detect::*_with_limits` functions take
   `&SourceLimits` as a required argument. `transaction::Commit` and
   `BatchCommit` replace their public fields with the triad accessors
   (`file()`, `into_file()`, `patch()`, `diagnostics()`, `into_parts()`,
   `value()`, `into_value()`), the facade's spelling of `package()` for a
   `File`.

### 2.6 Consumers, fidelity, performance

1. Every wave migrates, in the same commits as the API change, every
   consumer: litchi-ops (format modules, tests, fuzz), the facade
   (`crates/litchi/src`, examples, `crates/litchi/fuzz`), litchi-cli,
   pyo3-litchi (and `tests/run_tests.sh`), `tools/perf-baseline` (with its
   `Cargo.lock`), `tools/native-resave`, the crate's own examples, benches,
   tests and fuzz targets. No alias or deprecated forwarder remains.
2. Fidelity: byte-identical output for exact no-ops and for every edit the
   old path supported; the crate's round-trip, CRUD and corpus tests pass
   unchanged except where they named a removed API.
3. Performance: each wave re-measures its formats' routes (§5) under the
   goal's PERFORMANCE BAR (open 1.05 or +20 µs; single-op write 1.10; 1k-op
   batch 2.0), release build, pinned CPU, ABBA, and adds its rows to
   `tools/unified_bench_matrix.py`. Outputs moving to `&self` and edits
   moving off `&mut self` are the two risks: a second `save` of one package
   must cost no more than the first did before the change, and `edit_x()`
   must not copy part payloads.

## 3. Inventory by item

The complete per-item list, with the target of each row, is
[`results/0351/inventory.tsv`](results/0351/inventory.tsv). The package
mutators group into these target families (target from
`tools/api_debt_inventory.py`; the wave may merge small families into a
neighbour, e.g. `custom_props` into `metadata`).

### 3.1 DOCX `Package` (85) → families

| Target | Methods |
|---|---|
| outputs, `&self` | `save` `save_encrypted` `save_plain` `save_plain_with_durability` `save_reencrypted` `save_with_durability` `to_bytes` `to_encrypted` `to_plain_bytes` `to_plain_stream` `to_reencrypted` `write_to` |
| signature verbs → new `Package` | `resign` `resign_with` `sign` `sign_with` `unsign` |
| `alt_chunk` | `add_alt` `insert_alt` `move_alt` `remove_alt` `replace_alt` |
| `attached_template` | `remove_attached_template` `set_attached_template_uri` |
| `bibliography` | `add_bibliography_source` `remove_bibliography_source` `replace_bibliography_source` |
| `chart` (internal) | `store_chart_graph` |
| `conflict` | `apply_conflict_patch` `apply_conflicts` |
| `content_control` | `apply_content_control_patch` `apply_content_controls` |
| `custom_xml` | `add_custom_xml` `order_custom_xml` `remove_custom_xml` `replace_custom_xml` `set_custom_xml` |
| primary `apply` | `apply_durable_document_patch` |
| `document_variable` | `clear_document_variables` `remove_document_variable` `set_document_variable` |
| `fonts` | `clear_font_embedding` `put_fonts` `remove_fonts` `set_font_embedding` |
| `glossary` | `put_glossary` `put_glossary_graph` `remove_glossary` `take_glossary_graph` |
| history → `&self -> Result<Commit>` | `publish_document_commit_with_history` `redo_document` `undo_document` |
| `ink` | `apply_durable_ink_patch` `apply_durable_ink_patch_with_limits` `apply_ink_patch` `publish_ink_edit` |
| `mail_merge` | `clear_mail_merge` `set_mail_merge` `update_mail_merge` `update_mail_merge_recipients` |
| delete (0342 `metadata::Edit`) | `custom_props_mut` |
| `numbering` | `apply_numbering_patch` |
| raw OPC `Edit` | `edit_opc` |
| `ribbon` | `put_ribbon` `remove_ribbon` |
| `story_hyperlinks` | `replace_story_hyperlink_texts` |
| `styles_with_effects` | `apply_styles_with_effects_commit` `apply_styles_with_effects_patch` `put_styles_with_effects` `remove_styles_with_effects` |
| `task_panes` | `apply_durable_task_panes_patch` `apply_durable_task_panes_patch_with_limits` `apply_task_panes_patch` `put_task_panes` `put_task_panes_with_limits` `remove_task_panes` `remove_task_panes_with_limits` |
| `tracked_revision` | `apply_tracked_revision_patch` `apply_tracked_revisions` |
| `vba` | `clear_vba` `put_vba` `set_vba` `set_vba_with` |
| `web` | `put_web` `remove_web` |

DOCX `writer::MutableDocument` (44, reached from `Builder::document()`) and
its 202 draft methods (`MutableParagraph`, `MutableRun`, `MutableTable`,
`MutableRow`, `MutableStyle`, `Section`, …) follow §2.1 rule 7.

### 3.2 PPTX `Package` (77) → families

| Target | Methods |
|---|---|
| outputs, `&self` (limits from `OpenOptions`) | `save` `save_encrypted` `save_encrypted_with_limits` `save_plain` `save_plain_with_durability` `save_reencrypted` `save_reencrypted_with_limits` `save_with_durability` `to_bytes` `to_encrypted` `to_encrypted_bytes` `to_encrypted_with_limits` `to_plain_bytes` `to_reencrypted` `to_reencrypted_bytes` `to_reencrypted_with_limits` |
| signature verb | `unsign` |
| `change_tracking` | `apply_change_tracking_commit` `apply_change_tracking_patch` |
| `cross_slide_copy` | `apply_cross_slide_copy_patch` `apply_cross_slide_copy_plan` |
| `metadata` (custom properties) | `put_custom_props` `remove_custom_props` |
| `fonts` | `put_fonts` `remove_fonts` |
| `ink` | `apply_ink_actions_patch` |
| `master_layout` | `add_slide_layout` `add_slide_master` `remove_slide_layout` |
| `math_properties` | `put_math_properties` `remove_math_properties` |
| `model3d` | `put_model3d` `remove_model3d` |
| `notes` | `apply_notes_commit` `apply_notes_patch` `clear_notes` `put_notes` `remove_notes` |
| raw OPC `Edit` | `edit_opc` |
| internal | `store_placeholder_shape` `store_slide_sync` |
| `placeholder_type_extension` | `apply_placeholder_type_extension_patch` `apply_placeholder_type_extension_patch_with_policy` `apply_placeholder_type_extension_slot_patch` `apply_placeholder_type_extension_slot_patch_with_policy` |
| `read_only_recommended` (one spelling) | `apply_read_only_recommended_commit` `apply_read_only_recommended_patch` `apply_readonly_recommended_commit` `apply_readonly_recommended_patch` `put_read_only_recommended` `put_readonly_recommended` `remove_read_only_recommended` `remove_readonly_recommended` |
| `shape_classification` | `put_shape_classification` `remove_shape_classification` |
| `shape_design_element` | `put_shape_design_element` `remove_shape_design_element` |
| `shape_designer_properties` | `put_shape_designer_properties` `put_shape_designer_properties_with_limits` `remove_shape_designer_properties` `remove_shape_designer_properties_with_limits` |
| `shape_tags` | `put_shape_tags` `remove_shape_tags` |
| `slide_copy` | `apply_slide_copy_plan` |
| `slide_designer_tags` | `put_slide_designer_tags` `put_slide_designer_tags_with_limits` `remove_slide_designer_tags` `remove_slide_designer_tags_with_limits` |
| `slide_removal` | `apply_slide_removal_patch` `apply_slide_removal_plan` |
| `styles` | `put_styles` `remove_styles` |
| `tags` | `put_tags` `remove_tags` |
| `task_panes` | `apply_task_panes_patch` |
| `zooms` | `put_zooms` `remove_zooms` |

PPTX drafts: `writer::MutableSlide`, `MutableShape` and the presentation
writer's `add_slide`/`slide_mut` (39) follow rule 7.

### 3.3 XLSX `Package` (35) and XLSB `Workbook` (50)

| Owner | Target | Methods |
|---|---|---|
| XLSX | outputs `&self` | `save_encrypted` `save_encrypted_with_limits` `save_reencrypted` `save_reencrypted_with_limits` |
| XLSX | signature | `unsign` |
| XLSX | `edit_x(&self)` (was `&mut self`) | `edit_calculation_metadata` `edit_custom_data` `edit_data_model` `edit_page_breaks` `edit_slicers` `edit_smart_tags` `edit_surveys` `edit_task_panes` `edit_timelines` (and their `_with_limits`/`_with` forms → `edit_x_with_options`) |
| XLSX | `apply_x(&self)` | `apply_calculation_metadata_patch` `apply_custom_data_patch` `apply_data_model_patch` `apply_page_breaks_patch` `apply_surveys_patch` |
| XLSX | into the families above | `put_custom_data*` `remove_custom_data*` `rename_custom_data*` `put_data_model` `remove_data_model`; `put_custom_props` `remove_custom_props` → `metadata` |
| XLSB | signature | `resign` `resign_with` `sign` `sign_with` `unsign` |
| XLSB | `apply_x` | `apply_calculation_chain*` `apply_cell_watches` `apply_custom_data*` `apply_data_model*` `apply_sparklines` `apply_theme*` `apply_xml_maps*` |
| XLSB | `edit_x` families | `merged_cell_range` (8 `*_merged_cell_range*`), `connections`, `scenarios` (4), `slicer_caches`, `slicers`, `timeline_caches`, `timelines`, `ribbon`, `task_panes`, `vba` (4: `clear_vba` `put_vba` `set_vba` `set_vba_with`) |
| XLSB | raw OPC `Edit` | `edit_opc` |
| XLSB | delete | `Package::into_workbook` |

XLSB drafts: `writer::MutableWorksheet` and the workbook writer's
`get_worksheet_mut` (61) follow rule 7 (`xlsb::builder::Worksheet`).

### 3.4 Other owners

| Owner | `&mut self` | Target |
|---|---:|---|
| DOC `raw::Package` | 24 | lazy readers take `&self` (interior cache) or the type becomes `raw::Reader` (§2.1.6) |
| RTF `raw::Document` | 163 | raw mutable tree under a non-owner name (§2.1.6) |
| ODT `mutable::MutableDocument` | 136 | crate-private; verbs via `Edit`/`Builder` (§2.4.5) |
| ODT `generic::Package` | 12 | `rdfa`, `text_meta`, `xforms_model`, `variable_declaration_group`, `protection` families |
| ODT `odc::Document` | 11 | an `axis`/`series` chart family |
| ODS `MutableSpreadsheet`, ODI `FlatImageTransaction`/`FrameEditor`, ODT `flat` drafts | 54, 66, 13 | rule 7 and §2.5.5 |
| PPT `writer::Chart`, XLS `chart::Chart` | 3, 2 | builder drafts (rule 7) |

### 3.5 PPT scoped families and `rollback` (item 16)

The 14 PPT `Transaction` types (`FontTransaction`,
`PictureBulletTransaction` and `picture_bullets::Transaction`,
`broadcast`, `client_anchor`, `client_data`, `diagram`, `document_structure`,
`external_media`, `master_layout`, `package::property_set`, `slide_order`,
`text_edit::{Transaction, SourceTransaction}`) become `x::Edit`, and the 15
`rollback` methods (the ten transactions above plus `SlideSyncEditor`,
`diagram::SlideEditor`, `document_comparison::Editor`,
`master::metadata::Editor` and its `template::Editor`) are deleted.

Re-plumbing: every PPT `rollback(self) -> Snapshot` returns `self.source`,
the base the transaction took by value. The writers do not consume it:
outside tests the only callers are the two `master::metadata` delegations.
The edit entry therefore changes from consuming the base to holding a cheap
clone of it (or `&'a Snapshot`), so the caller keeps the base and dropping
the `Edit` is the rollback. Where a base clone would copy record payloads,
the base moves behind `Arc` first; the wave measures `edit()` cost against
the current consuming entry. The `Editor` types that are not families
(`SlideSyncEditor`, `document_comparison::Editor`, `master::metadata`
editors) become `edit_x()` families on the PPT `Package` or the relevant
snapshot, with `Edit::reset` where an in-place discard is used. The 16
`is_changed` and 19 commit `snapshot()` accessors go with them.

Shared legacy crates: `litchi-ole-common` (10 transactions with
`rollback`), `litchi-crypto` (`spaces::Transaction`) and `litchi-ograph`
(`rollback`, `UnsafeEdit`) are used by PPT, XLS and DOC and are migrated in
W5, after W3. W3 registers PPT bans only for names that no longer reach the
PPT surface through those crates; W5 registers the rest.

## 4. Glossary gate

### 4.1 Rule kinds added in F0

`tools/check_api_glossary.py` gains (documented in glossary §11, tested in
`tools/test_check_api_glossary.py`):

- **Affix bans**: a ban may carry `"match": "prefix"` or `"suffix"`. The
  affix never matches the bare affix itself.
- **`immutable_owners`**: fully qualified owner types; any public method on
  them taking `&mut self` fails, and a missing owner fails.
- **`family_methods: true`**: every registered family must also expose
  `Commit::{package, into_package, patch, diagnostics, into_parts}` without
  `snapshot`/`into_snapshot`/`is_changed`, `Diagnostics::changed`,
  `Edit::commit` without `rollback`/`is_changed`/`snapshot`, and
  `Patch::{is_noop, inverse}`. A member that is another crate's type is
  checked in that crate (rustdoc lists no dependency methods).

### 4.2 Registered in F0 (the surface already conforms)

- `immutable_owners`: the `Package` and projection of DOC, XLS, PPT, RTF,
  every ODF crate (with `odt::flat::Document`, `ods::FlatSpreadsheet`,
  `odp::FlatPresentation`, `odg::FlatDrawing`, `odc::FlatChart`,
  `odi::FlatImage`), Pages, Keynote and Numbers; `docx::Document`,
  `pptx::Presentation`, `xlsx::Workbook`, `xlsb::Package` — 56 owner
  entries across 29 profiles.
- `family_methods: true` on the 22 DOCX, PPTX, XLSX, DOC, XLS, PPT, RTF and
  ODF profiles that register families (their families already pass).
- 150 bans in the profiles whose crate has none of the name today:
  `UnsafeEdit`, `rollback`, `Transaction` (exact and suffix), `is_changed`,
  `into_snapshot`, the `Mutable` prefix and the `_with_limits` suffix.

### 4.3 Registered by each wave

| Wave | Adds to its profiles | Glossary cells it flips |
|---|---|---|
| W1 DOCX | `immutable_owners` `Package`; bans `UnsafeEdit`, `rollback`, `Transaction` (both), `is_changed`, `into_snapshot`, `Mutable` prefix, `_with_limits` suffix, `custom_props_mut`, `into_workbook`; every scoped family registered | §1 *Immutable owner snapshot* (docx), §2 *New snapshot accessor*, *Diagnostics*, *Split commit*, *Patch application*, *No-op* (DOCX scoped), *Mutable writer* (DOCX `UnsafeEdit`), §2a line W1 |
| W2 PPTX | `immutable_owners` `Package`; `UnsafeEdit`, `Mutable` prefix, `_with_limits`, `readonly_recommended` spellings; families registered | §1 (pptx), §2 rows (PPTX scoped), *Mutable writer* (PPTX), line W2 |
| W3 PPT | `rollback`, `Transaction` (both), `is_changed`, `_with_limits`, `get_*` getters; `family_methods` stays; families registered | *Rollback* (PPT), *Configured open* (PPT `raw::Reader`), line W3 |
| W4 XLSX/XLSB/common | `immutable_owners` `xlsx::Package`, `xlsb::Workbook`; `Transaction` (both), `is_changed`, `_with_limits`, `into_workbook`, `Mutable` prefix; `family_methods` on `xlsb-*`, `ooxml-common-*` | §1 (xlsb `into_workbook`), §2 rows (XLSX/XLSB scoped, `is_changed`), line W4 |
| W5 DOC/XLS/RTF | `immutable_owners` `doc::raw::*`, `rtf::raw::*` as renamed; `UnsafeEdit` (XLS), `Transaction`, `is_changed`, `into_snapshot`, `_with_limits`, `get_*`; PPT `rollback` if W3 left it | *Mutable writer* (XLS `UnsafeEdit`), line W5 |
| W6 ODF | `immutable_owners` for `flat::*` owners; `rollback`, `Transaction`, `Mutable` prefix, `_with_limits`, `into_snapshot`, `Limits` (flat crates, exact); prefixed flat names | §4 *Coverage policy* (flat ODF, ODC), *Mutable writer* (ODT `MutableDocument`), line W6 |
| W7 iWork | `family_methods` on `numbers-*`, `pages`, `keynote-*`; `_with_limits` | §2 *Split commit* (iWork), line W7 |
| W8 `File` | `_with_limits` in the `ops-*`/`facade-iwork` profiles (and an all-features `ops` profile if the gate's build budget allows) | §4 *Configured open*, line W8, the §2a closure |

## 5. Performance routes per wave

Matrix: [`results/0289/bench-matrix.md`](results/0289/bench-matrix.md),
regenerated with `python3 tools/unified_bench_matrix.py`. The latest series
per format are the baseline legs.

| Wave | Routes to re-measure | Latest evidence |
|---|---|---|
| W1 | DOCX open, `save`/`to_bytes` twice on one package, `set`/`add` single-op, 1k batch, metadata set, tail-append (perf-baseline `docx_*`) | 0345 `docx-p0-final-bench`, 0342, 0338, 0337 ABBA |
| W2 | PPTX open, save twice, `set`/`add`/`remove`/`move`, batch scaling, metadata | 0348 `pptx-p0-final-bench`, 0346, 0342, 0338 |
| W3 | PPT `set`/`add`, batch, chunked text publication | 0323 `ppt-elements-repeat`, 0264, 0337 ABBA |
| W4 | XLSX open/commit/save, cell set, batch, metadata; XLSB `set`/metadata/batch (perf-baseline `xlsx_cell_values_abba`, `xlsb_crud`) | 0347, 0338, 0336, 0330 ABBA; XLSB 0337, 0326, 0322, 0315 |
| W5 | DOC `set`/`get`/batch; XLS `set`/batch/open; RTF `set`/`add`/batch | DOC 0324, 0311; XLS 0327, 0316; RTF 0328, 0326; 0337 ABBA |
| W6 | ODT/ODS/ODP/ODG/ODC open, edit, batch, flat open | 0337 ABBA, 0328, 0316, 0310, 0291–0292 `odf_gaps_bench` |
| W7 | Keynote/Numbers/Pages open, `set`, batch | 0340 ABBA, 0347 `numbers-swap`, 0329 |
| W8 | batch budget (2.0), Python and CLI rows | 0348, Python/CLI rows of the matrix |

## 6. Foundation landed in F0

- `litchi_core::edit` (`crates/litchi-core/src/edit.rs`): `Commit`,
  `Diagnostics`, `Refusal`, `RefusalKind`, module docs with a compiled
  example, five unit tests.
- litchi-ops: `impl Classify for litchi_core::edit::{Refusal, RefusalKind}`
  (one code per kind; a future kind is `unsupported_capability` until
  classified) and the test `edit_refusals_classify_by_kind_alone`. No
  litchi-ops error code is added: the refusal kinds map onto existing codes.
- Gate: the three rule kinds of §4.1 with six new tests, and the §4.2
  registrations. `docs/API_GLOSSARY.md`: the D1–D6 cells are 🔴 again and
  point at §2a; §2a holds the contract rows and the per-wave status lines;
  §5 gains the *Edit refused* row; §11 documents the rule kinds.
- `tools/api_debt_inventory.py` and `results/0351/inventory.tsv`.

No public API of a format crate changed, so no consumer migrates in F0 and
no capability or registry golden changes.

## 7. Verification

On this branch (trunk `5ff605692e` plus the F0 changes):

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-core -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`; same for `litchi-core` | 0, 0 |
| `cargo test -p litchi-core --all-features` (249 unit, 57 doc, integration) | 0 |
| `cargo test -p litchi-ops --all-features --lib` (949) and `--doc` (213) | 0, 0 |
| capability/registry goldens, all features and all-but-`encryption` (`--test capabilities --test registry`) | 0, 0 (unchanged, no re-bless) |
| `cargo test -p litchi --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-core -p litchi-ops` | 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` | 0: 40 profiles, 2,479 banned names, 253 families, 56 immutable owners, family methods in 22 profiles |
| tool unit tests (the five required modules plus `test_check_api_glossary`, `test_unified_bench_matrix`; 1,304 tests) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` (31 ok) | 0 each |
| `check_crate_boundaries.py`, `non_iwork_gate.py verify` | 0, 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (298) | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 |

Negative check: with `immutable_owners: ["litchi_docx::Package"]` on the
all-features DOCX rustdoc the gate reports exactly the 85 mutators of 0343
D1, and an `UnsafeEdit` ban reports its variants.

Not run: the 162 litchi-ops integration test targets other than
`capabilities` and `registry`. F0 changes no route: the only litchi-ops
change is a classification impl that no leaf error reaches yet.

## 8. Performance

No hot path changed: `litchi_core::edit` is new and unused by any route,
and the litchi-ops classification runs only on error paths. No measurement
was taken.

## 9. Gaps and follow-ups

- Everything in §0 W1–W8 is open; F0 only makes it decidable.
- `family_methods` does not prove that `Commit::package()` returns the
  format's `Package` type or that `apply_x` exists; signatures remain Rust
  and CRUD test obligations, as for the six-name check.
- The gate's litchi-ops profiles build three feature subsets, not all
  features; the File-level `*_with_limits` ban (W8) needs an all-features
  ops profile or one per subset.
- The inventory's family names are a heuristic over method names; each wave
  confirms the family split against the code (for example `chart_graph` is
  an internal store, `custom_props` belongs to `metadata`).
