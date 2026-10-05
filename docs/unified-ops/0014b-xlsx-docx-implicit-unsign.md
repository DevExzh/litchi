# 0014b — No implicit `unsign()` in litchi-xlsx and litchi-docx (F1), and up-front XLSB task-pane and Ribbon refusals (F2)

- Status: implemented on `uo/0014b-xlsx-docx-unsign`, based on `uo/m0a` at
  `0b79b2f61a` (record [0013](0013-signature-policy.md) merged). Follow-ups
  F1 (the `litchi-xlsx` and `litchi-docx` part) and F2 of record 0013 §6. The
  `litchi-pptx` part of F1 is record 0014a.
- Date: 2026-10-05
- Scope: `crates/litchi-xlsx`, `crates/litchi-docx`, `crates/litchi-xlsb`
  (task panes and Ribbon only). No change to `litchi-ops`, `litchi-opc` or
  `litchi-ooxml-common`; no new dependency edge.
- Governing records: ADR 0006 ("Signed input is `Signed<T>`. Editing requires
  consuming it through explicit signature stripping or edit-and-resign"),
  ADR 0003 (exact no-ops stay exact), record 0013 (§1 policy, §2.1 the XLSB
  classification this record follows, §2.2 the call-site families, §6 F1/F2),
  `docs/UNIFIED_OPS_GOAL.md` (correctness first, typed refusals, exact no-ops;
  breaking changes are acceptable when documented).
- `performance_claim: none`. Nothing was timed.

Path prefixes: `X/` = `crates/litchi-xlsx/src/`, `D/` = `crates/litchi-docx/src/`,
`B/` = `crates/litchi-xlsb/src/`, `C/` = `crates/litchi-ooxml-common/src/`.
Line numbers are at this record's commit unless a cell says "base"
(`0b79b2f61a`); the first column of every classification table is the base
line of the removed `unsign()` call.

## 1. The rule applied

Record 0013 §1: nothing strips a signature as a side effect of an edit. Every
implicit `unsign()` call site is one of:

- **(a)** it follows a refusal and only normalizes OPC policy flags on a
  package that is already unsigned. Keep it and say why.
- **(b)** it strips silently on a changed edit. Refuse with the crate's typed
  signed error, before anything is mutated, unless the caller removed the
  signature explicitly first.
- **(c)** it runs on an exact no-op. The no-op must not strip and must stay
  byte-identical.

The refusal is one helper per crate, the same test as XLSB's
`B/signature_policy.rs:21`: `is_signed() || requires_signature_edit_policy()`.
The second term catches a signature graph deleted part by part (generic
removal is not stripping, ADR 0006), which the old trailing `unsign()`
authorized silently.

| Crate | Helper | Typed error |
|---|---|---|
| `litchi-xlsx` | `X/signature_policy.rs:24` (new) | `Error::Signed`. Connections, XML Maps and volatile dependencies return `litchi_core::sheet::Result`, so their refusal is a boxed `litchi_xlsx::Error::Signed` (downcast it) |
| `litchi-docx` | `D/signature_policy.rs:24` (new) | `Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)`. DOCX has no `Signed` variant; this is the OPC refusal its source-backed editors and PPTX's commit (record 0013) already use, and `litchi-ops` classifies it `document_protected` with the `Signed::unsign()` suggestion. `Error::UnsafeEdit`, the other DOCX refusal, classifies as `unsupported_capability` and was not used |
| `litchi-xlsb` | `B/signature_policy.rs:21` (existing) | `PackageError::Signed` |

Every check sits after the call's own no-op detection and before its first
mutation of the caller's package (or of the staged clone a transaction
publishes), so a refused call leaves an owned package exact: it still
republishes its signed source byte for byte. The trailing `unsign()` calls are
gone; on an unsigned package they only reset policy flags that no longer
matter, so removing them changes nothing there.

The explicit authorization, documented on every affected public item:

- `litchi_opc::OpcPackage::unsign` for the raw `&mut OpcPackage` APIs;
- `litchi_xlsx::Package::unsign(&mut self) -> Result<&mut Self>` (new,
  `X/package.rs:260`) for `Package::edit_slicers`, `edit_timelines` and
  `edit_smart_tags`; it refuses a package opened with a password (encryption
  policy) and leaves a package without a signature or a signature-policy
  requirement untouched, exact source included, like PPTX's `Package::unsign`;
- `litchi_xlsx::Workbook::unsigned` (record 0013) for workbook snapshots;
- `litchi_docx::Package::unsign` (existing; its documentation now says it is
  the authorization);
- `litchi_xlsb::Workbook::unsign` (existing).

What counts as an exact no-op here is what each function already treats as
"nothing to do" (removing something absent, an unchanged transaction, an empty
patch, an equal value where the family compares values), plus the (c) paths
below, which now detect that their output equals the stored bytes. Raw stores
and update closures that always rewrite a part were not given new semantic
no-op detection: they are refused on a signed package even when the rewrite
would reproduce the stored content, which is XLSB's behaviour after record
0013 (its §3 item 5). §7 lists that as a gap.

## 2. Classification

Method: every site was read in context with every public path that reaches
it. A temporary probe test (built, run, deleted, as record 0013 did) confirmed
the "before" behaviour on the producer-signed fixtures: `connections::store_in_package`
on both signed XLSX fixtures returned `Ok` and the published bytes reopened
unsigned; `Package::put_web` and `Package::put_ribbon` on both signed DOCX
fixtures returned `Ok` and left the package unsigned; the XLSB rows are §3.
The "after" behaviour is pinned by the tests named in §5.

Counts by grep of `unsign(` (call sites; the two test-function names that
contain `unsign` are not calls):

| Crate | Call sites | (a) | (b) | (c) | Explicit API | Tests |
|---|---|---|---|---|---|---|
| `litchi-xlsx` | 49 | 0 | 36 | 6 | 1 (`Workbook::unsigned`) | 6 |
| `litchi-docx` | 13 | 0 | 11 | 0 | 1 (`Package::unsign`) | 1 |

No site of either crate followed a refusal, so there is no (a) row: every
implicit `unsign()` either stripped a changed package or ran unconditionally.

### 2.1 litchi-xlsx

Reached-from names are public paths; `crud` is the private
`X/slicer_cache/crud.rs`, reached through `slicer::Transaction` and
`timeline::Transaction` (`Package::edit_slicers`, `Package::edit_timelines`).

| Base site | Function | Reached from | Class | After | Exact no-op on a signed package |
|---|---|---|---|---|---|
| `X/revisions/package.rs:212` | `store_workbook_revisions_inner` | `revisions::store_workbook_revisions`; `replace_workbook_revisions` from `revisions::Transaction::commit` and `revisions::Patch::apply` | b | refused before the first part is added (`X/revisions/package.rs:178`); `commit` and `apply` refuse before cloning (`X/revisions/transaction.rs:286`, `X/revisions/patch.rs:65`) | unchanged commit, empty patch (existing) |
| `X/revisions/package.rs:304` | `remove_workbook_revisions_inner` | `revisions::remove_workbook_revisions`, `replace_workbook_revisions` | b | refused (`:297`) | no owner: `false` (existing) |
| `X/revisions/package.rs:483` | `update_in_place` | `replace_workbook_revisions` when the owner keeps its topology | b (the strip was already guarded by a change test) | an unchanged owner returns before any mutation (it used to take `get_part_mut`, revoking the exact source); a change is refused (`:428`) | unchanged owner |
| `X/workbook/comments/package.rs:102` | `store_on_worksheet` | `workbook::comments::{store_on_worksheet, replace_on_worksheet}`, `comments::Transaction::commit`, `comments::Patch::apply` | b | refused before the part is written (`:80`); `commit` `X/workbook/comments/transaction.rs:236`, `apply` `X/workbook/comments/patch.rs:70` | unchanged commit, empty patch |
| `X/workbook/comments/package.rs:140` | `remove_from_worksheet` | same | b | refused (`:145`) | no comments: `false` |
| `X/connections/package.rs:158` | `store_in_package_with_query_table_validator` | `connections::store_in_package` | b | refused in both branches (`:140`, `:145`) | none (a store always rewrites) |
| `X/connections/package.rs:207` | `remove_from_package` | `connections::remove_from_package` | b | refused (`:213`) | no connections: `false` |
| `X/connections/package.rs:1326` | `apply_connections` | `connections::Transaction::commit` | b | `(None, None)` publishes nothing (unreachable from `commit`, which returns first; it used to `unsign` anyway); otherwise refused (`:1312`); `commit` refuses before cloning (`:1143`) | unchanged commit |
| `X/slicer_cache/crud.rs:70` | `add_slicer`, existing part | `slicer::Transaction::add_view` | b | refused (`:67`) | — |
| `X/slicer_cache/crud.rs:81` | `add_slicer`, new part | same | b | `store_slicer_part` refuses (`X/slicer_cache/views.rs:426`) | — |
| `X/slicer_cache/crud.rs:128` | `update_slicer` | `update_view`, `replace_view` | b | refused (`:126`) | absent name: `false`; a closure that changes nothing still rewrites, so it is refused |
| `X/slicer_cache/crud.rs:182` | `remove_slicer` | `remove_view` | b | refused (`:166`) | absent name: `false` |
| `X/slicer_cache/crud.rs:215` | `reorder_slicers` | `reorder_views` | c | publishes nothing when every rewritten part equals its stored bytes, always for a worksheet without slicers (`rewrites_nothing`, `:211`, `:724`); a real reorder is refused (`:214`) | worksheet without slicers; current order of library-written parts |
| `X/slicer_cache/crud.rs:237` | `add_slicer_cache` | `add_cache` | b | `store_slicer_cache` refuses (`X/slicer_cache/package.rs:183`) | — |
| `X/slicer_cache/crud.rs:260` | `update_slicer_cache` | `update_cache`, `replace_cache` | b | refused (`:260`) | absent name: `false` |
| `X/slicer_cache/crud.rs:315` | `remove_slicer_cache` | `remove_cache` | b | refused (`:307`) | absent name: `false` |
| `X/slicer_cache/crud.rs:344` | `reorder_slicer_caches` | `reorder_caches` | c | publishes nothing when the rewritten workbook extension equals the stored bytes (`:344`); otherwise refused (`:347`) | current order of a library-written extension |
| `X/slicer_cache/crud.rs:366` | `add_timeline_cache`, first cache | `timeline::Transaction::add_cache` | b | `store_timeline_caches` refuses (`X/timelines/package.rs:197`) | — |
| `X/slicer_cache/crud.rs:400` | `add_timeline_cache`, further cache | same | b | refused (`:388`) | — |
| `X/slicer_cache/crud.rs:424` | `update_timeline_cache` | `update_cache`, `replace_cache` | b | refused (`:426`) | absent name: `false` |
| `X/slicer_cache/crud.rs:482` | `remove_timeline_cache` | `remove_cache` | b | refused (`:476`) | absent name: `false` |
| `X/slicer_cache/crud.rs:511` | `reorder_timeline_caches` | `reorder_caches` | c | as `reorder_slicer_caches` (`:513`, `:516`) | only a rewrite equal to the stored bytes; the extension `store_timeline_caches` writes is reserialized differently by the reorder, so the current order of one written by this library is still a change and is refused (F1c) |
| `X/slicer_cache/crud.rs:561` | `add_timeline`, existing part | `add_view` | b | refused (`:566`) | — |
| `X/slicer_cache/crud.rs:573` | `add_timeline`, new part | same | b | `store_worksheet_timelines` refuses (`X/timelines/package.rs:415`) | — |
| `X/slicer_cache/crud.rs:614` | `update_timeline` | `update_view`, `replace_view` | b | refused (`:618`) | absent name: `false` |
| `X/slicer_cache/crud.rs:677` | `remove_timeline` | `remove_view` | b | refused in both branches (`:667`, `:681`) | absent name: `false` |
| `X/slicer_cache/crud.rs:709` | `reorder_timelines` | `reorder_views` | c | publishes nothing when the rewritten part equals the stored bytes (`:714`); otherwise refused (`:717`) | current order of a library-written part; a worksheet without timelines already returned early |
| `X/active_x/package/mod.rs:114` | `remove_from_worksheet` | `active_x::remove_from_worksheet`, `replace_on_worksheet` | b | refused (`:127`); the `unsign()` it replaces ran before the mutation | no controls: `false` |
| `X/active_x/package/transaction.rs:181` | `install_graph` | `active_x::store_on_worksheet`, `replace_on_worksheet` | b | refused (`:181`) | — (an empty set is invalid input) |
| `X/xml_maps/snapshot.rs:196` | `Snapshot::restore_into` | `xml_maps::Patch::apply` | b | refused before the first mutation (`:118`); `apply` refuses before cloning (`X/xml_maps/patch.rs:63`) | empty patch |
| `X/xml_maps/package.rs:101` | `store_xml_in_package` | `xml_maps::store_in_package`, `xml_maps::Transaction::commit` | b | refused in both branches (`:73`, `:92`); `commit` refuses before cloning (`X/xml_maps/transaction.rs:233`) | an equal value and conformance publish nothing (existing guard in `store_in_package`); unchanged commit |
| `X/xml_maps/package.rs:125` | `remove_from_package` | `xml_maps::remove_from_package` | b | refused (`:128`) | no XML Maps: `false` |
| `X/slicer_cache/package.rs:194` | `store_slicer_cache_in_place` | `slicer::store_cache`, `add_slicer_cache` | b | refused (`:183`) | — |
| `X/volatile_dependencies/package.rs:81` | `store_in_package` | `volatile_dependencies::store_in_package` | b | refused in both branches (`:54`, `:71`) | none (a store always rewrites) |
| `X/volatile_dependencies/package.rs:105` | `remove_from_package` | `volatile_dependencies::remove_from_package` | b | refused (`:107`) | no part: `false` |
| `X/smart_tags/package.rs:27` | `store` | `smart_tags::Transaction::{replace, set, remove, clear}` (`Package::edit_smart_tags`) | c | publishes nothing when the worksheet bytes do not change, e.g. replacing absent smart tags with an empty collection (`:32`); otherwise refused (`:37`). It used to rewrite the identical bytes and strip | `replace` with an empty collection where none exists; `remove` of an absent cell and `clear` of an empty collection return before `store` (existing) |
| `X/named_sheet_view/package.rs:90` | `store_worksheet_named_sheet_views` | same | b | refused in both branches (`:83`, `:89`) | none (a store always rewrites) |
| `X/named_sheet_view/package.rs:117` | `remove_worksheet_named_sheet_views` | same | b | refused (`:119`) | no part: `false` |
| `X/query_table/package.rs:107` | `add_worksheet_query_table` | same | b | refused (`:97`) | — |
| `X/query_table/package.rs:133` | `replace_worksheet_query_table` | `replace_…`, `update_worksheet_query_table` | b | refused (`:138`) | none (a replace always rewrites) |
| `X/query_table/package.rs:170` | `remove_worksheet_query_table` | same | b | refused (`:182`) | absent relationship: `false` |
| `X/query_table/package.rs:234` | `reorder_worksheet_query_tables` | same | c | publishes nothing when every relationship already has its deterministic ID, type and target in the requested order, always for a worksheet without query tables (`relationships_already_ordered`, `:244`, `:271`); otherwise refused (`:247`). It used to remove and re-add the same relationships and strip | worksheet without query tables; current canonical order |

Explicit, kept: `X/workbook/model.rs:1342` is `Workbook::unsigned` itself
(record 0013). `X/survey.rs:5850`, `:5867`, `X/workbook/data_model/package.rs:2286`
and `X/custom_data/package.rs:2806`, `:2820`, `:2835` are tests that remove the
signature explicitly after asserting a refusal.

### 2.2 litchi-docx

| Base site | Function | Reached from | Class | After | Exact no-op on a signed package |
|---|---|---|---|---|---|
| `D/package/package/merge.rs:196` | `Package::set_mail_merge` | `set_mail_merge`, `update_mail_merge`, `update_mail_merge_recipients` | b | refused after staging, before the first part is installed (`:180`) | none (setting always republishes) |
| `D/package/package/merge.rs:291` | `Package::clear_mail_merge` | same | b | refused (`:301`) | no mail merge: `false` (existing) |
| `D/glossary/graph.rs:440` | `put_graph` | `glossary::put`, `glossary::raw::put`, `Package::put_glossary`, `Package::put_glossary_graph`, `glossary::Transaction::commit`, `glossary::Patch::apply` | b | refused before the candidate is built (`:449`); `commit` and `apply` refuse before cloning (`D/glossary/transaction.rs:468`, `:606`) | identical graph: `false` (existing, documented) |
| `D/glossary/graph.rs:501` | `remove_graph` | `glossary::remove`, `glossary::raw::remove`, `Package::remove_glossary`, `Package::take_glossary_graph`, `commit`, `apply` | b | refused (`:510`) | no glossary: `None` (existing) |
| `D/web/package.rs:90` | `web::apply_patch` | `apply_patch`, `apply_commit` | b | refused before staging (`:89`) | a patch that leaves the bytes unchanged (existing) |
| `D/web/package.rs:153` | `web::put`, existing part | `web::put`, `Package::put_web` | b | refused (`:162`) | equal settings and conformance: `false` (existing) |
| `D/web/package.rs:173` | `web::put`, new part | same | b | refused (`:178`) | — |
| `D/web/package.rs:209` | `web::remove` | `web::remove`, `Package::remove_web` | b | refused (`:217`) | no part: `false` (existing) |
| `D/font/package.rs:308` | `font::put` | `font::put`, `Package::put_fonts`, managed font embedding at save | b | refused at the start of the mutation phase (`:314`) | equal table: `false`; empty table without a table: `false` (existing) |
| `D/font/package.rs:449` | `remove_graph` | `font::remove`, `Package::remove_fonts`, `font::put` of an empty table | b | refused (`:459`) | no table: `false` (existing) |
| `D/package/model.rs:536` | `ensure_word_font_settings` | managed font embedding at save (`Package::set_font_embedding`, feature `automatic-fonts`) | b | refused (`:536`); the save rolls back | settings already carrying the flags: `false` (existing) |

Explicit, kept: base `D/package/package/access.rs:111` (now `:116`) is
`Package::unsign` itself;
`D/content_control/package.rs:1874` is a test that removes the signature
explicitly after asserting a refusal.

Also a DOCX family, though the call site is in the shared crate:
`Package::put_ribbon` and `Package::remove_ribbon` (base `D/package/model.rs:733`,
`:743`) call `litchi_ooxml_common::ribbon::{put, remove}`, which remove the
signature themselves (`C/ribbon/package.rs:113`, `:148`, `:191`; the probe
confirmed the strip on both signed DOCX fixtures). The facade now refuses a
change to a signed package before the shared owner is reached
(`D/package/model.rs:748`, `:771`): the version and bytes the family already
stores, and removing an absent family, stay exact no-ops. The check reads the
Ribbon slots only when the package is signed, so an unsigned package takes
the old path unchanged. Same fix as F2 (§3).

### 2.3 Late refusals in the same families, now up front

These were not `unsign()` sites: they mutated a signed package in memory
without removing the signature, so the OPC writer refused it only at save,
after the caller's package had changed. Each now refuses before mutating, so
the family behaves one way.

| Site | Before | After |
|---|---|---|
| `connections::Patch::apply` → `restore_snapshot` | mutated, refused at save | refused before cloning (`X/connections/package.rs:1220`) |
| `slicer::package::store_part` → `store_slicer_part` | mutated, refused at save | refused before the first part is added (`X/slicer_cache/views.rs:426`) |
| `timeline::package::store_caches` → `store_timeline_caches` | same | `X/timelines/package.rs:197` |
| `timeline::package::store_part` → `store_worksheet_timelines` | same | `X/timelines/package.rs:415` |

## 3. F2: XLSB task panes and Ribbon

Record 0013 §2.1 listed `put_task_panes`, `remove_task_panes`, `put_ribbon`
and `remove_ribbon` as mutating in place and refused only at save ("read"
evidence). The probe showed otherwise, so this corrects that row:

| Method | Before (probe, `Simple.xlsb` signed in-test) | After |
|---|---|---|
| `Workbook::put_task_panes` | refused before mutation by the shared web owner's patch, but as `PackageError::Common(litchi_ooxml_common::Error::Opc(SignedSourceRequiresExplicitPolicy))`, which `litchi-ops` classifies `corrupt_file`; equal panes were already an exact no-op | the planned patch is checked: a non-empty patch over a signed workbook is refused with `PackageError::Signed` before it is applied (`B/workbook/package.rs:616`, check `:624`); equal panes still publish nothing |
| `Workbook::remove_task_panes` | same `Common` refusal; absent panes `false` | `PackageError::Signed` (`:636`, check `:648`); absent panes still `false` |
| `Workbook::put_ribbon` | **silent strip** through `C/ribbon/package.rs:113`/`:148`: `Ok`, the workbook unsigned, the save published without a signature | refused with `PackageError::Signed` unless the family already stores the same version and bytes (`:677`, check `:683`); the identical put stays an exact no-op |
| `Workbook::remove_ribbon` | **silent strip** (`C/ribbon/package.rs:191`) | refused with `PackageError::Signed` when the family is present (`:696`, check `:700`); an absent family stays `false` |

`web::put`/`web::remove` were replaced by their own `plan_put`/`plan_remove`
followed by `Patch::apply`, which is what they run, so an unsigned workbook
takes the same path. The Ribbon slot is read only for a signed workbook.

## 4. Behaviour changes

1. **litchi-xlsx**: every (b) and (c) family of §2.1 refuses a change to a
   signed package with `Error::Signed` (boxed for connections, XML Maps and
   volatile dependencies) and leaves the package as it was. They used to
   remove the signature and publish. Remove it first with
   `OpcPackage::unsign`, `Package::unsign` or `Workbook::unsigned`.
2. **litchi-xlsx**: a package whose signature graph was deleted part by part
   (the OPC edit-policy requirement survives) is refused the same way; the
   trailing `unsign()` used to authorize it silently.
3. **litchi-xlsx**: a store into a signed package is refused even when it
   would rewrite the stored content (connections, comments, volatile
   dependencies, named sheet views, query-table replace, slicer and timeline
   updates whose closure changes nothing). XML Maps keep their existing
   equal-value no-op.
4. **litchi-xlsx**, all packages: a slicer, slicer-cache, timeline or
   timeline-cache reorder whose rewrite equals the stored bytes, a query-table
   reorder whose relationships are already the ones it would write, and a
   smart-tag store that leaves the worksheet bytes unchanged publish nothing.
   On an unsigned package the content is the same as before; the package now
   keeps its exact source (and the query-table relationships keep their order
   in the `.rels` part) instead of being rewritten.
5. **litchi-xlsx**: `connections::Patch::apply`, `slicer::package::store_part`
   and `timeline::package::{store_caches, store_part}` refuse a signed package
   before mutating it instead of at save (§2.3).
6. **litchi-xlsx**: new `Package::unsign(&mut self) -> Result<&mut Self>`.
7. **litchi-docx**: every family of §2.2 refuses a change to a signed package
   with `Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)` and leaves it
   as it was; it used to remove the signature. `Package::put_ribbon` and
   `remove_ribbon` refuse the same way. A managed save with automatic font
   embedding that would add fonts or change the font settings of a signed
   document fails with that error and rolls back; it used to publish the
   document without its signature. Remove it first with `Package::unsign` (or
   `OpcPackage::unsign` on the raw APIs).
8. **litchi-docx**: `font::put`'s documentation said it "invalidates any
   package signatures immediately before the mutation phase"; it now says the
   opposite. The web-settings documentation was corrected the same way.
9. **litchi-xlsb** (F2): task panes refuse with `PackageError::Signed` instead
   of `PackageError::Common(..)`; the Ribbon refuses instead of stripping
   (§3).
10. Unsigned packages: no other change. The checks pass, the removed trailing
    `unsign()` only reset flags that do not matter on an unsigned package, and
    the F2 and DOCX Ribbon checks run only for a signed package.

## 5. Tests

New:

- `crates/litchi-xlsx/tests/package_signature_policy.rs` (16 tests): per
  family (connections, query tables, XML Maps, volatile dependencies, classic
  comments, revisions, named sheet views, `ActiveX`, smart tags, slicers and
  timelines) the signed refusal with the package left byte-identical, the
  exact no-ops byte-identical, and success after the explicit removal with no
  signature part left. Stores run on both producer-signed fixtures
  (`hello-world-signed.xlsx`, `ms-office-2010-signed.xlsx`); removals, patches
  and the (c) reorders run on a package that holds the family's content and is
  signed in the test (`sign`, the XLSB test's construction). Patches whose
  snapshot binds the package root (connections, XML Maps) are applied to the
  signed graph with its signature parts deleted by hand, which still refuses
  (§4 item 2). Plus `Package::unsign` on both fixtures.
- `crates/litchi-xlsx/src/package.rs`: a doctest for `Package::unsign`.
- `crates/litchi-docx/tests/package_signature_policy.rs` (7 tests): web
  settings (facade and raw patch), font table, glossary (facade, raw graph,
  transaction and patch), mail merge, Ribbon, on all three signed fixtures
  (`hello-world-signed.docx`, `hello-world-signed-twice.docx`,
  `ms-office-2010-signed.docx`) and on documents signed in the test.
- `crates/litchi-docx/src/package/model.rs`:
  `font_settings_tests::font_settings_refuse_a_signed_package_until_it_is_unsigned`
  (feature `automatic-fonts`): the managed font settings refuse both DOCX
  fixtures byte-identically, succeed after `unsign`, and an already-flagged
  signed package is an exact no-op.
- `crates/litchi-xlsb/tests/signature_policy.rs`: module `task_panes_and_ribbon`
  (2 tests) for F2.

Rewritten because they pinned the old stripping:

- `crates/litchi-docx/tests/glossary_authoring.rs`:
  `graph_noop_preserves_signature_and_real_change_unsigns` became
  `graph_noop_preserves_signature_and_real_change_requires_explicit_unsign`:
  the changed graph is refused and the signature kept, then `Package::unsign`
  authorizes it.
- `crates/litchi-docx/src/glossary/tests.rs`:
  `transaction_noop_preserves_signature_and_patch_rejects_stale_source`: the
  changed commit over the signed clone is now asserted refused; the patch it
  goes on to check comes from an explicitly unsigned copy.

No `litchi-xlsx` or `litchi-xlsb` test pinned a strip. The shared
`litchi-ooxml-common` Ribbon tests still pin its own strip; that owner is
unchanged (§7).

## 6. Gates

On this branch, with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0` and `3rdparty`
symlinked from the main checkout for the test runs (removed afterwards):

- `cargo fmt --check -p litchi-xlsx -p litchi-docx -p litchi-xlsb`: clean.
- `cargo clippy -p <crate> --all-targets --all-features -- -D warnings`, run
  once per crate for `litchi-xlsx`, `litchi-docx` and `litchi-xlsb`: clean.
- `cargo test -p <crate> --all-features --no-fail-fast`, all passing:
  `litchi-xlsx` 2,156 (record 0013: 2,139); `litchi-docx` 2,030, 32 ignored
  (2,022); `litchi-xlsb` 1,041, 12 ignored (1,039). No evidence directory had
  to be added to the sparse checkout.
- `RUSTDOCFLAGS=-Dwarnings cargo doc -p <crate> --no-deps --all-features` for
  the three crates: clean.
- `cargo test -p litchi-ops --all-features signature`: 26 passed, among them
  record 0013's audit `tests::signature::every_signed_input_follows_the_explicit_signature_policy`
  and `an_unlocked_signed_package_is_not_unsigned_into_the_clear`.
- `cargo check -p litchi-ops -p litchi --all-features`: clean.
- The two pre-commit checks of `AGENTS.md`: both print nothing.

## 7. Follow-ups

- F1b: the shared-owner strippers in `litchi-ooxml-common`, outside this
  record's grep: `C/ribbon/package.rs:113`, `:148`, `:191` (DOCX and XLSB now
  refuse before reaching them; a direct caller still gets a strip),
  `C/custom/package.rs:97`, `:121`, `:126` (custom properties),
  `C/properties/write.rs:68`, `:95`, `:115` (core properties),
  `C/custom_xml/package.rs:184`, `:492` (Custom XML; DOCX's data stores
  already refuse first), `C/vba.rs:211`, `:267` (VBA). They are reached from
  at least `litchi_xlsx::Package::{put_custom_props, remove_custom_props}`
  and the DOCX save path (edited custom or core properties), and from PPTX and
  XLSB, so one record across the four leaves (with 0014a's PPTX owner) should
  classify them site by site. Not classified or fixed here.
- F1c: semantic no-op detection for the raw stores and update closures of §4
  item 3, and for the reorders whose rewrite does not reproduce the stored
  bytes (a timeline-cache reorder into the current order reserializes the
  extension `store_timeline_caches` wrote), so an unchanged request keeps a
  signed package exact (ADR 0003), as F8 of record 0013 asks for the eager
  XLSX transaction.
- M2: the xlsx connections, XML Maps and volatile-dependencies families
  return boxed errors, so their refusal must be downcast to
  `litchi_xlsx::Error::Signed` before it can be classified.
- Incidental, not a signature issue and not fixed:
  `add_timeline_cache` validates the existing timeline views against the new
  cache alone (base `X/slicer_cache/crud.rs:363`), so adding a second timeline
  cache to a workbook that already has a timeline view fails with "references
  missing cache".

## 8. Not claimed

- Nothing was timed. The F2 and DOCX Ribbon checks add one Ribbon read for a
  signed package only; not measured.
- No signature is verified: signatures are detected, and removed on request.
- The "before" column rests on the probe for the rows §2 names and on reading
  the code for the rest; the probe test was deleted.
- §2.3 lists the late refusals of the families touched here, not every late
  refusal in `litchi-xlsx`.
- The unified layer does not reach these families yet (record 0013 §2.2), so
  nothing in `litchi-ops` changed.
