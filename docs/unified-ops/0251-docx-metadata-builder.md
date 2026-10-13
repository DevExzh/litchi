# 0251 — DOCX metadata transaction, Builder and stream output (ADR 0034 W2 residue)

Plan step 5 (docx part) of [0248](0248-coordinator-restart-plan.md) §6 Phase B.
Branch `uo2/p7a-docx`, cut from trunk `908f9cccdd`.

## Contract

### Core-properties family

`litchi_docx::metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` is
the core-properties transaction family of ADR 0034 §7, with the package entry
points `Package::{metadata, edit_metadata, apply_metadata}`:

- `metadata()` borrows the typed core properties; `None` is a package without a
  core-properties part, so absence survives a read.
- `edit_metadata()` stages `set`, `update` and `clear` over a cheap package
  clone. Dropping the edit discards it.
- `Edit::commit()` publishes through the shared OOXML codec
  (`litchi_ooxml_common::properties::{write, clear}`) into a clone of the
  package graph. The new package's property slot is reloaded from the graph,
  so it carries no dirty state, and the source package is unchanged.
- An equal value is an exact no-op: the commit returns a clone that shares every
  part, signature included, and `diagnostics().changed()` is false.
- A changed value on a signed package is `metadata::Error::Signed` at commit
  time; before this change it was a refusal at save.
- A text property over 1 MiB (the shared codec's per-property bound) is
  `Error::Limit { kind: LimitKind::PropertyTextBytes, .. }` before any staging.
- `Patch` carries the base value and the target value. `is_noop()` and
  `inverse()` are the canonical introspection. `apply_metadata(&patch)` refuses a
  package whose current value differs from the base with
  `Error::PatchConflict`.

The family is a thin DOCX wrapper over what `litchi-ooxml-common` has today. The
value type is the shared `properties::Props`, re-exported as `metadata::Props`.
When step 6 (branch `uo2/p6-ooxml-meta`) merges `Props` into `Metadata`, only
the re-export, the `admit` field list and the call sites of
`Edit::{set, update}` change; the family shape does not. This change does not
edit `litchi-ooxml-common/src/properties`.

**Update after merging 0253.** Step 6 merged as record 0253, and this
branch merged it. The value type is now `litchi_core::Metadata`,
re-exported as `metadata::Metadata`, and `admit` checks `author`. As
predicted, nothing else in the family changed. The docx family still wraps
the core codec only. Staging application and custom properties through the
shared `litchi_ooxml_common::properties` engine, as pptx, xlsx and xlsb do,
is a follow-up.

`Package::{props, props_mut, put_props, clear_props}` are deleted. The package
slot still exists internally, so the internal retry tests keep using it
directly.

### `document_mut` deleted; authoring moves to `Builder`

`Package::document_mut()` is deleted. Fresh authoring uses `litchi_docx::Builder`:

- `Builder::{new, new_template}`.
- `document()` returns the legacy `MutableDocument` authoring model.
- `package()` gives the package-level authoring verbs (glossary entries,
  alternative chunks, custom XML) during construction.
- `font_embedding(mode)`, with the `automatic-fonts` feature.
- `build()` materializes the authored document into the package graph without
  a ZIP round trip (a new `Package::finish_authoring`, sharing the save path's
  materializer), embeds fonts when requested, flushes staged property facades,
  and returns a package with no authoring state.

This follows the PPTX precedent of [0127](0127-pptx-canonical-constructor.md)
and ADR 0034 §4 ("option-bearing authoring uses one Builder per format").

**Deviation from the step text.** The plan says "move its authoring verbs into
`Edit`". ADR 0034 §9 bans hydrating an opened package into a lossy writer, and
`MutableDocument` is that writer. Its verbs therefore live on the fresh-content
`Builder`, not on the opened-package `document::Edit`.

The package-level legacy mutators that hydrate internally are not migrated here:

- `add_alt`, `insert_alt`, `replace_alt`, `remove_alt` and `move_alt`;
- watermark, text box, OLE and VML authoring;
- content-control package edits;
- VBA and revision authoring.

They reach the writer through the crate-private `authoring_document()`, and
their migration remains a follow-up. ADR 0034 §9 is therefore still partial for
DOCX, and the glossary row says so.

**Capability lost on opened packages.** An opened document could previously
change through hydration: adding headings, adding hyperlink paragraphs, and
setting or removing protection. It now changes through `Package::edit()`
alone. Protection editing of an opened package has no replacement in this
change. Hyperlink insertion is in step 7a's scope.

### Stream output

The source-backed tail-append `Plan::write_to_stream` and `Commit::write_to_stream`
(both `tail_append` and `tail_append_stream`) are renamed `write_to`. OPC's own
`write_to_stream` (`PackageWriter`, `SourceBackedPackage` splices) is outside
this change.

### Consumers migrated

- `litchi-ops`:
  - the eager metadata read now uses `metadata()`;
  - the DOCX creator applies initial metadata through `edit_metadata()`;
  - `metadata::Error` maps through a new `Source::DocxMetadata` source (signed →
    `document_protected`, limit → `limit_exceeded`, patch conflict →
    `io_error`).
- Facade examples (`litchi`): `core_props_office`, `office_crud_demo`,
  `comprehensive_docx_test`, `docx_template_test`, `docx_with_fonts`,
  `minimal_font_test` and `read_office_files`.
- Five DOCX examples, `README.md` and `docs/OFFICE_API_GUIDE.md`.
- `tools/perf-baseline` (compile-level):
  - fixture generators use `Builder`;
  - the tail-append routes use `write_to`.
  - **Behaviour change.** The frozen `ordinary_save` DOCX edit route used
    `document_mut().add_paragraph_with_text`. It now appends the same marker
    through `Package::edit().insert_paragraph` and commit, and its description
    string says so. Results recorded before this change measured the hydrating
    writer and are not comparable with new runs of that route.

Tests that asserted hydration of an opened package, or dirty-writer refusals on
opened packages, are deleted. These behaviours are no longer reachable:

- `byte_order_marked_parts`: the mutable-edit differential and the
  protection-edit differential;
- `web_extensions/durable`: the dirty-document refusal;
- `ink_inventory/removal`: the pending-writer refusal;
- `real_file_preservation`: the changed-relationships hydration test.

Fresh-package dirty-writer refusals keep their coverage through
`Builder::package()`. The opened-package edit tests moved to `Package::edit()`:

- `paragraph_index_reuse`
- `owned_commit`
- `glossary_authoring`
- `custom_properties`
- `conflict_revisions`

Signature-policy tests now assert the commit-time `metadata::Error::Signed`
refusal, then the explicit `unsign()` path.

## Glossary

- `tools/api_glossary.json`: both DOCX profiles ban `document_mut`, `props`,
  `props_mut`, `put_props` and `clear_props` on `litchi_docx::Package`, and
  `write_to_stream` crate-wide. Both register the `litchi_docx::metadata`
  family.
- `docs/API_GLOSSARY.md`: the owner, mutable-writer, stream-output,
  metadata-read and metadata-mutation rows are flipped for DOCX, and a 0251
  section is appended.

## Verification

Environment: the worktree target, `CARGO_INCREMENTAL=0`, no debug info,
toolchain 1.95.0.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-docx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-docx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features` | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-ops` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile docx-minimal` / `--profile docx-full` | 0 / 0 (46 bans, 1 family each) |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `cargo check --manifest-path tools/perf-baseline/Cargo.toml --all-targets` | 0 |
| `cargo check -p litchi-docx --no-default-features --all-targets`; `cargo check -p litchi-ops --no-default-features --features docx` | 0 |

`cargo check --workspace --all-targets --all-features` was also run. Its only
errors were pre-existing and outside this change:

- missing purged evidence files that `soapberry-zip`, `litchi-opc` and
  `litchi-ods` test targets `include_bytes!`;
- the iWork chart examples.

## Not claimed

- No performance claim: this step changes no measured hot path.
- No pptx/xlsx metadata family. No `app.xml` or custom-property transaction:
  `custom_props_mut` remains, as glossary row "Custom properties" records.
- No migration of the remaining package-level hydrating mutators.
