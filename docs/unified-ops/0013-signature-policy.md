# 0013 — The explicit signature policy (M0a step S6b)

- Status: implemented on `uo/m0a-s6b-signed`, based on `uo/m0a` at
  `d3a1812182`. Step S6b of record
  [0003](0003-m0a-migration-plan.md) §4 (policy in §2.6).
- Date: 2026-10-05
- Scope: `crates/litchi-ops` (the policy, the editor builders, an audit test),
  `crates/litchi-pptx` and `crates/litchi-xlsb` (the silent strippers),
  `crates/litchi-xlsx` (an explicit removal for the eager snapshot), and
  read-only accessors in `crates/litchi-docx`, `crates/litchi-pptx`,
  `crates/litchi-xlsb` and `crates/litchi-ppt`. One new dependency edge,
  `litchi-ops -> litchi-sign` (ADR 0024 amendment of 2026-10-05).
- Governing records: ADR 0006 ("Signed input is `Signed<T>`. Editing requires
  consuming it through explicit signature stripping or edit-and-resign"), ADR
  0003 (exact no-ops stay exact), ADR 0033 §4, record 0002 cross-cutting
  finding 4 (the silent strippers), record 0009 D1 (write verbs take `&self`).
- `performance_claim: none`. Nothing was timed.

Path prefixes: `S/` = `crates/litchi-ops/src/`, `P/` = `crates/litchi-pptx/src/`,
`B/` = `crates/litchi-xlsb/src/`, `X/` = `crates/litchi-xlsx/src/`,
`D/` = `crates/litchi-docx/src/`. Line numbers are at this record's commit
unless a row says "base" (`d3a1812182`).

## 1. The policy

A signed input is visible in the type, readable, and not editable until the
caller consumes the signature explicitly. Nothing strips a signature as a side
effect of an edit.

| State | How a `File` gets there | Reads | Editor builders |
|---|---|---|---|
| unsigned | the opener found no signature | yes | build as before |
| signed | the opener found OPC signature infrastructure or a compound-file `_signatures`/`_xmlsignatures` entry; the file is `File::Signed`, and the file it wraps (`Signed::file`) is marked signed too | yes, through the wrapped file | refuse: `document_protected`, suggestion "remove the signature explicitly with `Signed::unsign()` before editing; ..." |
| removal authorized | `Signed::unsign(self) -> File` | yes, the same content | build over a copy whose signature the leaf removed explicitly; every write from that editor publishes without signature parts |

API (all in `litchi-ops`):

- `Signed::unsign(self) -> File` (`S/signed.rs:164`). Consuming, infallible,
  O(1): it reads and rewrites nothing. The returned file carries the internal
  "removal authorized" disposition. Other clones of the same `Signed`, and
  `Signed::file()`, stay signed: the authorization belongs to the returned
  handle only. `#[must_use]`.
- `File::is_signed(&self) -> bool` (`S/file.rs:251`): `true` for a
  `File::Signed` and for the file it wraps; `false` for the file `unsign`
  returns, although its bytes keep the signature until its first write.
- The disposition lives in `S/format/signature.rs`: `Signature::{Unsigned,
  Signed, RemovalAuthorized}` (`:24`) and `Opened<S>` (`:42`), one handle's
  view of the state of one open. Every tier 1 and tier 2 handle is now
  `Arc<Opened<State>>` instead of `Arc<State>`, so a signed file and the file
  `unsign` returns share the owner, caches and retained source and differ only
  in the disposition. `Deref` reaches the state, so the read code did not
  change; each handle is still one pointer and `File` is still 16 bytes (the
  compile-time assertions hold). The cost is one small allocation per open
  and one pointer hop per read; not measured.
- `Signed::new` marks the wrapped file signed (`S/signed.rs:102`);
  `File::with_signature` (`S/file.rs:295`) re-marks a tier 1 or tier 2
  variant.

What "removal authorized" does, per editor builder. Each copy is built from
the handle's own retained source, never from the file system again:

| Handle | Builder | Signed | After `unsign` (the editor returned) |
|---|---|---|---|
| `Docx` | `edit_document` → `DocxEditor` (`S/format/docx.rs:407`) | refused | `DocxEditor::Unsigned(Box<litchi_docx::Package>)`: an owning package over the exact input bytes (source owner) or a clone of the eager owner's OPC graph, then `Package::unsign` |
| `Pptx` | `into_editor` (`S/format/pptx.rs:518`) | refused (handle returned in `EditorRefused`) | `PptxEditor::Unsigned(Box<pptx::Package>)`: same construction, then the new `Package::unsign` |
| `Xlsx` | `into_editor` (`S/format/xlsx.rs:409`) | refused | `XlsxEditor::Eager(workbook)` where `workbook` is the new `Workbook::unsigned` copy (exact input bytes for the source owner) |
| `Xlsb` | `editor` (`S/format/xlsb.rs:345`) | refused | `XlsbEditor::Materialized(workbook)`: exact input bytes (source owner) or the eager owner's OPC graph, then `Workbook::unsign` |
| `Doc` | `body_text_editor` (`S/format/doc.rs:188`) | refused | `SourceSnapshot` over the retained compound file with its signatures cleared by `litchi_sign::cfb::Editor::clear` |
| `Xls` | `cell_values_editor` (`S/format/xls.rs:247`) | refused | `cell_values::Snapshot` over the cleared compound file |
| `Ppt` | `text_editor` (`S/format/ppt.rs:242`) | refused | `text_edit::SourceSnapshot` over the cleared compound file |

Details:

- The signature check precedes the `retain_source(false)` check, so a signed
  handle always reports `document_protected`.
- The exact input bytes come from `SourceBackedPackage::source_artifact()`
  (`S/format/signature.rs:107`), the promotion record 0009 D4 recommends: the
  copy keeps every ZIP member the edit and the signature removal do not touch.
  It is opened under the open's OOXML read limits, which the OOXML handles now
  keep (`ooxml_limits`).
- An unlocked encrypted DOCX or PPTX refuses with `unsupported_capability`:
  a copy of its graph would drop the encryption profile the eager owner keeps,
  and the leaves refuse to edit a package they decrypted anyway. An unlocked
  encrypted XLSX refuses through the leaf's own encryption policy
  (`Workbook::unsigned`). XLSB has no decryption. A legacy encrypted file's
  retained source is still encrypted, and its leaf editor refuses it.
- The editor-builder signatures changed (they are crate-internal and only
  tests call them; M2 consumes them): `Docx::edit_document` returns
  `DocxEditor { Edit(Box<Edit>), Unsigned(..) }`; the DOC, XLS and PPT builders
  return `litchi_ops::Result` instead of the leaf error, so the policy refusal
  is a unified error. Three leaf errors gained typed sources and classifiers:
  `Source::DocBodyText` (`litchi_doc::body_text::source::Error`),
  `Source::PptTextEdit` (`litchi_ppt::text_edit::Error`) and `Source::Sign`
  (`litchi_sign::Error`), plus a classifier for `litchi_cfb::OverlayError`.

`resign` (re-signing with caller keys) is deferred, as record 0003 §2.6
planned it (P1, behind a `sign` feature). Re-signing is edit-and-resign: the
signature has to be computed over the bytes the write publishes, and the
unified layer publishes nothing until M2's commit adapter and `save` exist. It
also needs a signer input that meets ADR 0006's credential rules (non-`Clone`,
redacted, zeroized) and that `OpenOptions` deliberately cannot carry, plus the
verification statuses (integrity, coverage, validity, trust, revocation,
time) that P1 owns. The leaves keep their own signing APIs
(`OpcPackage::sign`/`resign`, `litchi_sign::cfb::Editor::add`/`resign`).

## 2. Leaf audit

Every leaf commit or publish entry the unified editor builders reach, plus the
tier 3 strippers record 0002 found. "Before" is the base commit; "after" is
this change. Evidence: **test** = asserted by a test added or kept in this
change; **probe** = a temporary probe test built, run and deleted during this
change (as record 0009 §2.4 did), on the producer-signed fixtures and the
in-test signed XLSB and compound files; **read** = read from the code, not run.

### 2.1 Tier 1 and tier 2: entries the builders reach

| Leaf entry (reached from) | Before: changed commit on signed input | After | Exact no-op on signed input | Evidence |
|---|---|---|---|---|
| DOCX `Package::edit_document` → `apply_document_patch` (`Docx`, eager owner) | refused: `Error::UnsafeEdit` "changed signed content requires explicit Package::unsign" (`D/package/package/access.rs:234-240`) | unchanged; the builder refuses first | byte-identical | probe; test (no-op) |
| DOCX `source_backed::Package::edit_document` → `publish_document_commit_to_stream` (`Docx`, source owner) | refused: OPC `SignedSourceRequiresExplicitPolicy` | unchanged | byte-identical | probe; test (no-op) |
| PPTX `opened::Transaction::commit` → `Package::apply_opened_presentation_commit` → `to_bytes` (`Pptx`, eager owner) | **silent strip**: `commit` called `working.unsign()` (base `P/opened/transaction.rs:1262`); `apply_opened_presentation_commit` then removed the signature parts from the package in memory by generic part removal, and `to_bytes`/`save` refused it with `SignedSourceRequiresExplicitPolicy` | refused at `commit` with `Error::Opc(SignedSourceRequiresExplicitPolicy)` (`P/opened/transaction.rs:1265`, `refuse_signed_change` `:1357`); authorize with the new `Package::unsign` (`P/package/model.rs:1879`) | byte-identical | test (`litchi-pptx/tests/opened_signature_policy.rs`, `opened::tests::changed_commit_refuses_a_signed_package_until_it_is_unsigned`); probe (before) |
| PPTX `SourceBackedPresentationEditor::edit_slide` → `publish_slide_commit_to_stream` (`Pptx`, source owner) | refused: OPC `SignedSourceRequiresExplicitPolicy` | unchanged | byte-identical | probe; test (no-op) |
| XLSX `Workbook::edit` / `Edit::commit` / `Workbook::apply` (`Xlsx`, eager owner) | refused: `Error::Signed` (`X/workbook/edit/codec.rs:99`, called at `X/workbook/edit/semantic/transaction.rs:337`, `:1502`) | unchanged; new explicit removal `Workbook::unsigned` (`X/workbook/model.rs:1335`) | **refused** (not exact): the transaction refuses a signed workbook even for a no-op; it never strips | probe; test |
| XLSX per-feature `SourceBackedEditor`s over the OPC overlay publisher (`Xlsx`, source owner, `XlsxEditor::Source`) | refused by the OPC overlay publisher | unchanged | byte-identical (defined-names editor) | test (no-op); read |
| XLSB `Workbook::apply_cell_values[_by_name]`, `cell_values::workbook::apply*` | **silent strip** (base `B/cell_values/workbook.rs:167`) | `Error::Signed` (`B/cell_values/workbook.rs:164`) | byte-identical | test (`litchi-xlsb/tests/signature_policy.rs`); probe (before) |
| XLSB workbook structure: `WorkbookEdit` rename, transfer, authored strings and formulas, images, drawing transfer; `WorkbookPatch::apply` | **silent strip** (base `B/cell_values/root.rs:1086`, `:1257`, `:1331`, `:1513`; `B/cell_values/drawing_transfer.rs:284`) | `Error::Signed` (`B/cell_values/root.rs:559`, `:1087`, `:1132`, `:1516`; `B/cell_values/drawing_transfer.rs:238`); a rename to the current name is now an exact no-op | byte-identical | test (rename, empty commit); read (others) |
| XLSB `Workbook::apply_sparklines` | **silent strip** (base `B/sparkline/workbook.rs:70`) | `Error::Signed` (`B/sparkline/workbook.rs:66`) | unchanged path (no change is published) | test (`sparkline::worksheet_tests::workbook_api_refuses_a_change_to_a_signed_workbook`) |
| XLSB `Workbook::apply_cell_watches` | **silent strip** (base `B/cell_watches/workbook.rs:134`) | `Error::Signed` (`B/cell_watches/workbook.rs:129`) | unchanged path | test |
| XLSB `Workbook::edit_opc` and the typed setters built on it: slicer and timeline caches and views, scenarios, merged cells | **silent strip**, unconditionally, even for a closure that changed nothing (base `B/workbook/package.rs:701`) | signature state preserved; a changed candidate is refused with `SignedSourceRequiresExplicitPolicy` unless the closure removes the signature explicitly (`OpcPackage::unsign`) (`B/workbook/package.rs:715`), as DOCX's `edit_opc` does | byte-identical (no longer stripped) | test |
| XLSB `Workbook::apply_xml_maps[_patch]` (through `edit_opc`) | **silent strip** (the transaction also plans its target without signatures) | `Error::Signed` before staging (`B/xml_maps/workbook.rs:44`); plan after `Workbook::unsign` | unchanged (an empty patch publishes nothing) | test (`xml_maps_public.rs`, three tests updated) |
| XLSB `Workbook::set_vba`/`put_vba`/`clear_vba` | **silent strip** of the package signature (base `B/host/vba_project.rs:282`, `:308`) | `Error::Signed` (`B/host/vba_project.rs:255`, `:307`); VBA project signatures are still removed with the project, as before | not applicable | read (feature `vba-inspection`; compiled and clippy-checked) |
| XLSB `Workbook::set_connections`/`remove_connections` | **silent strip** (base `B/host/connections/package.rs:140`, `:170`) | `Error::Signed` (same lines) | removing an absent graph is a no-op | test |
| XLSB raw threaded-comment helpers (`comments::threaded::store_graph`, `store_on_worksheet`, `remove_from_worksheet`, `remove_graph`, `Patch::apply`) | **silent strip** (base `B/comments/threaded/package/mod.rs:223`, `:500`, `:569`) | `Error::Signed` (`:218`, `:463`, `:519`) | `remove_graph` with nothing to remove stays exact | test |
| XLSB raw slicer and timeline helpers (`slicer::package::store_caches`/`store_views`, `timeline::package::...`) and the `Package`-level slicer and timeline transactions | **silent strip** (base `B/slicer/package.rs:470`, `:523`, `:625`, `:660`; `B/timeline/package.rs:417`, `:461`, `:551`, `:584`; `B/slicer/transaction.rs:271`; `B/timeline/transaction.rs:265`) | `Error::Signed` (`B/slicer/package.rs:457`, `:587`; `B/timeline/package.rs:406`, `:519`); the transactions no longer unsign after storing | unchanged (unchanged content is not stored) | test (`slicer_timeline.rs`, one test updated; `set_slicer_caches` in `signature_policy.rs`) |
| XLSB theme, theme owner, data model, calculation chain, custom data, volatile dependencies | already refused (e.g. `B/theme.rs:458`, `B/data_model/patch.rs:66`, `B/calculation_chain/patch.rs:76`, `B/volatile_dependencies/patch.rs:68`, `B/custom_data/package.rs:943`); volatile dependencies keep their explicit `without_signatures` | unchanged | unchanged | read |
| XLSB `put_task_panes`, `remove_task_panes`, `put_ribbon`, `remove_ribbon` | mutate the package in place without removing the signature; the OPC writer refuses the changed signed package at `save` | unchanged (late refusal) | unchanged | read |
| DOC `body_text::source::SourceSnapshot` (`Doc`) | refused at admission: `Refusal::Signed` "signed DOC sources are unsupported" (`litchi-doc/src/body_text/source.rs:1881`) | unchanged; the builder refuses first | no signed no-op path (refused at admission) | probe; test |
| XLS `cell_values::Snapshot::from_shared_ole_file` / `from_bytes_with_limits` (`Xls`) | refused at admission: protected container (`litchi-ole-common/src/protection.rs:34`) | unchanged | no signed no-op path | probe; test |
| PPT `text_edit::SourceSnapshot::from_shared` / `open` (`Ppt`) | refused at admission: protected container in the overlay validation (`litchi-ole-common/src/protection.rs:116`) | unchanged | no signed no-op path | probe; test |

Not reached by any builder, recorded for completeness: DOC eager
`body_text::Snapshot` refuses a signed container at admission (probe); PPT
`slide_order::Snapshot` opens a signed container but refuses every edit
(`Refusal::UnsupportedSource`, `litchi-ppt/src/slide_order.rs:1051`; probe for
the open, read for the refusal).

### 2.2 Implicit `unsign()` call sites outside the builders' reach

The builders hand out either a transaction detached from the owning package,
a source-backed editor, or (after `unsign`) an owning copy that is already
unsigned, so the leaves' package-level families never see a signed package
through the unified layer. The leaves still have implicit `unsign()` calls in
those families, about 60 in `litchi-pptx` (comments, modern comments, notes,
tags, change tracking, shapes and themes, presentation properties, embedded
objects, fonts, `Package::edit_opc`), about 10 in `litchi-docx` (mail merge,
glossary, web settings, font table, font settings) and about 45 in
`litchi-xlsx` (comments, connections, revisions, ActiveX, named sheet views,
slicer and timeline CRUD on `xlsx::Package`, volatile dependencies, XML maps,
smart tags, query tables). Some follow a refusal and only normalize the OPC
policy flags; the others strip silently for direct leaf callers. They need the
classification §2.1 gave XLSB; that is follow-up F1, not done here.

### 2.3 Tier 3 (record only)

The common mechanism: `litchi-odf-common`'s rebuild copies every source member
except the signature owners (`crates/litchi-odf-common/src/core/writer.rs:1621`,
in `copy_source_files_from_except` `:1597`), so every family that rebuilds a
signed package without its own signed-source check drops the signature.

| Stripper | Where | Proposed fix |
|---|---|---|
| ODF formula `Edit::commit` | `crates/litchi-odf-formula/src/facade/mod.rs:744` → `Package::replace_content` (`crates/litchi-odf-formula/src/package/mod.rs:156`) → `rebuild_package` (`crates/litchi-odf-common/src/package/edit.rs:635`); no signed check (read) | refuse a changed commit when the source holds a signature owner (`is_signature_owner_path`), as ODS's DDE and sheet-metadata families do, and add an explicit policy mirroring ODG's `SecurityWritePolicy::{Refuse, RemoveSignatures}` |
| ODP in-place annotation mutators `Presentation::add_annotation`/`replace_annotation`/`remove_annotation` | `crates/litchi-odp/src/package/presentation.rs:369-415` → `annotation::codec` (`crates/litchi-odp/src/annotation/codec.rs:131`, `:160`, `:179`) → `annotation::package::rebuild` (`crates/litchi-odp/src/annotation/package.rs:7`); no signed check (read) | the same refusal, or move the mutators onto the ODP transaction, whose `SecurityPolicy` already refuses signed packages (record 0002) |
| ODP in-place RDF mutators | `crates/litchi-odp/src/package/presentation.rs:629-715` (record 0002 §ODP; not re-read site by site) | as above |
| ODS legacy `MutableSpreadsheet`/`Spreadsheet` mutators | record 0002 §ODS (code-path inference, not re-run); the rebuilds without an in-function check include `crates/litchi-ods/src/document.rs:1736`, `:3082`, `:3187` and `crates/litchi-ods/src/charts/package.rs:11`, while DDE metadata, sheet metadata and database ranges already refuse (`crates/litchi-ods/src/facade/mod.rs:545`, `:694`; `crates/litchi-ods/src/database_range/package.rs:20`) | classify each mutator the way §2.1 classifies XLSB and refuse the stripping ones until an explicit `SignatureWritePolicy` is passed |

None of these is small and isolated: each needs a refusal plus an explicit
authorization API, and no tier 3 format is editable through the unified layer
yet. Not implemented (F3). The unified opener also does not detect ODF
signatures, so ODT, ODS and ODP never open as `Signed`; that must change
before their write handlers land (F4).

## 3. Behaviour changes

1. **PPTX** `opened::Transaction::commit` refuses a changed commit over a signed
   package with `Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)`.
   Before, it removed the signature from the commit's snapshot, and publishing
   the commit removed it from the package in memory, although the package then
   refused to save. New `Package::unsign(&mut self) -> Result<&mut Self>`
   removes the signature explicitly; it refuses a signed package opened with a
   password (encryption policy), and leaves a package without a signature or
   a signature-policy requirement untouched, exact source included.
2. **XLSB**: every changed publication listed in §2.1 as a silent strip now
   refuses a signed package with `PackageError::Signed` (or, through
   `edit_opc`, `PackageError::Opc(SignedSourceRequiresExplicitPolicy)`).
   `Workbook::unsign()` before the edit is the explicit authorization;
   `OpcPackage::unsign` inside an `edit_opc` closure also is.
3. **XLSB** `Workbook::edit_opc` no longer removes signatures before running
   the closure: a closure that changes nothing keeps the source bytes, and a
   changed signed candidate is refused unless the closure unsigns. Its
   documentation said it dropped signatures; it now says the opposite.
4. **XLSB** XML Maps: a changed patch refuses a signed workbook before
   staging; plan the edit after `Workbook::unsign`. The inverse still never
   restores a signature.
5. **XLSB** slicer and timeline caches and views (raw helpers, the
   `Package`-level transactions and the `Workbook` setters): a store into a
   signed package is refused, even when the store would rewrite identical
   content. `remove_*` of an absent part stays a no-op.
6. **XLSB** `WorkbookEdit::rename_sheet` to the sheet's current name publishes
   nothing (it used to reparse and, on a signed workbook, strip).
7. **XLSB** `opc_package_mut` (crate-internal) no longer unsigns.
8. **XLSX** new `Workbook::unsigned(&self) -> Result<Workbook>`: a copy with
   every signature removed explicitly; refuses a workbook opened with a
   password; returns an unsigned workbook as itself. No existing behaviour
   changed.
9. **litchi-ops**: the editor builders refuse a signed file with
   `document_protected` (they built editors before; the leaves then refused or
   stripped). Their return types changed (§1). `File::is_signed` and
   `Signed::unsign` are new public API. `Signed::file()` now returns a file
   whose builders refuse. The suggestion of every signature refusal, the
   leaves' included, now names `Signed::unsign()`. The DOC, XLS and PPT leaf
   errors map to typed codes instead of being returned raw.
10. New read-only accessors: `source_backed_package()` on
    `litchi_docx::source_backed::Package`, `litchi_pptx::SourceBackedPresentation`
    and `litchi_xlsb::SourceBackedWorkbook` (as XLSX's L-XLSX-1), and
    `litchi_ppt::SourceBackedPackage::shared_ole_file()`.

New leaf tests: `litchi-pptx/tests/opened_signature_policy.rs` (4),
`litchi-xlsb/tests/signature_policy.rs` (8), the XLSB sparkline test above,
`litchi-xlsx/tests/workbook_unsigned.rs` (2). New litchi-ops tests: the audit
of §4 (2) and classifier tests for the three new leaf error sources (3).

Tests changed to the new behaviour: `litchi-pptx`
`opened::tests::commit_rehashes_a_package_whose_unsign_strips_signature_bytes`
became `changed_commit_refuses_a_signed_package_until_it_is_unsigned`;
`litchi-xlsb` `slicer_timeline::patch_binds_root_relationships_and_inverse_restores_signature_state`
became `..._needs_an_explicit_unsign`; three `xml_maps_public` tests now
refuse first and edit after `Workbook::unsign`; the litchi-ops DOCX handle
tests match `DocxEditor::Edit`.

## 4. The audit test

`S/tests/signature.rs::every_signed_input_follows_the_explicit_signature_policy`,
over every signed fixture a `File` opens in the build: the three DOCX, two
PPTX and two XLSX fixtures S6 found under `test-data/poi/test-data/xmldsign/`;
`Simple.xlsb` signed in-test (an OPC origin, signature part and relationships,
as `litchi-xlsb`'s own tests build them); and `documentProperties.doc`,
`FormulaSheetRange.xls` and an authored text-box PPT, each signed in-test with
both a legacy `_signatures` stream and an `_xmlsignatures` storage. The
compound-file bases are files the narrow source-backed editors admit; S6's
`FancyFoot.doc` and `SampleShow.ppt` are refused by those editors whether or
not they are signed. Each input opens on four routes, by bytes and by path,
each source-backed and eager (the compound-file routes ignore the mode): 44
route cases in an all-features build. Each route asserts:

1. it opens as `File::Signed` with the right format and origin, and reads;
2. the builder of the wrapped file refuses with `document_protected`, the
   format, and a suggestion naming `Signed::unsign()`; `Signed::file()`
   reports `is_signed()`;
3. `signed.clone().unsign()` is not signed, reads the same text, and its
   builder's editor publishes a changed commit (a new paragraph, text box,
   cell, or a same-width text or number replacement) whose output has no
   `_xmlsignatures` member, no OPC signature graph, no `_signatures` stream and
   no `_xmlsignatures` storage, reopens unsigned, and shows the edit; the
   `Signed` it was cloned from still refuses;
4. per input, an exact no-op through the leaf entry the builder reaches
   reproduces the signed input byte for byte (DOCX and PPTX on both owners,
   XLSX source-backed, XLSB eager); the eager XLSX transaction and the three
   compound-file editors refuse the signed input instead, which the test pins.

`an_unlocked_signed_package_is_not_unsigned_into_the_clear` encrypts the
signed DOCX, unlocks it, and checks that the unsigned file's builder refuses
with `unsupported_capability` rather than produce a plaintext copy.

## 5. Gates

On this branch, with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0` and `3rdparty`
symlinked from the main checkout for the test runs (removed afterwards):

- `cargo fmt --check -p litchi-ops -p litchi-xlsb -p litchi-pptx -p litchi-xlsx -p litchi-docx -p litchi-ppt`: clean.
- `cargo clippy --all-targets --all-features -- -D warnings` on those six
  crates: clean. litchi-ops also with no default features and with each of
  `doc docx ppt pptx xls xlsx xlsb rtf odt ods odp pages encryption` alone and
  `docx,encryption`, `pptx,encryption`, `doc,encryption`, `doc,ppt,xls`,
  `pptx,xlsx`, `xlsx,xlsb`, `docx,doc`: clean.
- `cargo test --all-features --no-fail-fast`, all passing: litchi-ops 366
  unit tests and 54 doctests; litchi-xlsb 1,039 (12 ignored); litchi-pptx
  1,314 (3 ignored); litchi-xlsx 2,139; litchi-docx 2,022 (32 ignored);
  litchi-ppt 1,349 (11 ignored).
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the six
  crates: clean.
- `cargo check -p litchi --all-features`: clean.
- `python3 tools/check_crate_boundaries.py`: 66 workspace packages, 265
  internal dependency declarations; `python3 tools/non_iwork_gate.py verify`:
  47 bulk roots; `python3 -m unittest tools/test_check_crate_boundaries.py`
  (1,130 tests; the pinned policy-edge count moves from 263 to 264) and
  `python3 -m unittest tools.test_non_iwork_gate` (59 tests): pass.

## 6. Follow-ups

- F1: classify the implicit `unsign()` sites of §2.2 in `litchi-pptx`,
  `litchi-docx` and `litchi-xlsx` and refuse the stripping ones, before M2
  exposes those families.
- F2: XLSB task panes and ribbon refuse only at save (§2.1); refuse up front.
- F3: the tier 3 strippers of §2.3.
- F4: detect ODF signatures in the unified opener (an `Odf` origin) before the
  tier 3 write handlers land.
- F5: a file `unsign` returned and never edited still holds the signed source,
  because the removal happens when its first editor is built. M2 decides what
  its `save` writes: the source bytes, or the unsigned copy.
- F6: `resign` (P1, §1).
- F7: the copies `unsign` builds are not charged to an execution context; a
  managed open's exact-source copy and reopen sit outside its budget (record
  0009's open point on promotions).
- F8: the eager XLSX transaction refuses exact no-ops on signed workbooks;
  ADR 0003 would let them through.
- F9: litchi-ppt's text editor reports its protected-container refusal as an
  OLE `InvalidFormat`, which classifies as `corrupt_file`; a signed handle no
  longer reaches it, but a DRM-protected one would.

## 7. Not claimed

- Nothing was timed. The extra allocation per open and pointer hop per read
  that `Opened` adds are not measured.
- No signature is verified anywhere: signatures are detected at open, and
  removed on request.
- The probe evidence of §2 is a temporary test that was deleted; its
  conclusions are pinned by the tests named in §2 except where a row says
  "read".
- §2.2 counts call sites by grep; they were not classified site by site.
- The tier 3 rows are read from code (and record 0002), not run.
- Fidelity of the unsigned copy is argued from the exact-source promotion, not
  byte-compared member by member.
