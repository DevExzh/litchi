# ADR 0034: Unified cross-format API glossary and default policies

- Status: Accepted for this branch by owner instruction on 2026-10-08
- Date: 2026-10-08

## Adoption and numbering

The owner requested this design for the unified-operations refactor on
2026-10-08. The supplied base-checkout draft was numbered 0033; that number
already belongs to the accepted unified-operations ADR in this branch. This
record therefore takes 0034, preserving the existing ADR and its historical
citations. [Record 0016](../unified-ops/0016-rebase-and-glossary-adoption.md)
records rollout adjustments for already-landed M0a and independently owned
iWork crates. Acceptance specifies the target contract; it does not claim
that the migration waves or enforcement checker have landed.

This branch briefly tracked a pointer file at the draft's path,
`docs/adr/0033-unified-api-glossary-and-default-policies.md`. It duplicated
ADR number 0033 and collided with the untracked draft in the base checkout,
so [record 0334](../unified-ops/0334-facade-features-and-batch-fuzz.md)
removed it. This file is the only copy of the glossary ADR in this branch.

## Context

The workspace grew format by format, and each family landed the ADR 0003 /
0004 contracts with a different local vocabulary. The public surface today
exhibits the same concept under several names and several default policies:

- Top-level owner types: `Package` (docx, pptx, numbers, pages, keynote, doc),
  `Workbook` (xlsx, xlsb, xls), `Document` (odt), `Spreadsheet` (ods),
  `Presentation` (odp, ppt). The semantic snapshot is called `Document` in all
  three iWork crates, including Keynote, where the value is a presentation.
- Creation: `Package::new()` (docx, pptx), `Workbook::new()` and
  `Workbook::create()` as duplicates (xlsx), `Builder` (odt, ods, odp),
  `WorkbookWriter` (xlsb), `Writer` (doc), `*DocumentBuilder` (litchi-iwa).
- Edit entry: `edit()` (xlsx, odt, keynote), `edit_document()` (docx),
  `opened_presentation_transaction()` (pptx, whose transaction type is named
  `Transaction`, not `Edit`), per-feature `edit_*()` (numbers).
- Commit shape: docx and the iWork families carry `Diagnostics`; the xlsx and
  pptx top-level commits do not. Publication verbs are `into_workbook()`
  (xlsx), `into_package()` (numbers), `publish_document_commit()` (docx), and
  `apply_opened_presentation_commit()` (pptx).
- Byte sources: `from_bytes(Vec<u8>)` (xlsx, odt, odp), `from_bytes(&[u8])`
  (pptx, numbers), `from_slice(&[u8])` (xlsx), `from_shared_bytes(Arc)`
  (numbers), plus `from_reader` and `from_ole_file`.
- Configuration: `ReadLimits` (OPC), `Limits` + `OpenOptions` +
  `PackageOpenOptions` (doc, which has two overlapping option types),
  `SourcePackageLimits` (odt), `DocumentSourceLimits` (iWork), and semantic
  limits named `DocumentLimits` (numbers), `SemanticLimits` (pages), and
  `DocumentSemanticLimits` (keynote).
- Errors: one god-enum per crate in OOXML (`xlsx::Error` spans style mapping
  to ZIP failures) versus per-focused-module errors in iWork; the same
  wrong-family failure is `UnexpectedFormat { detected }` (facade),
  `NotPages`/`NotKeynote`/`NotNumbers` (iWork), or a string inside
  `InvalidFormat` (odt). `sheet::Workbook` returns `Box<dyn Error>` while the
  sibling facades return typed errors, and `map_ooxml_error` flattens leaf
  errors into strings.
- Metadata: two near-identical types, `litchi_core::Metadata` and
  `ooxml_common::Props` (`author` vs `creator`), with typed package-level
  access only in docx.
- Selectors: the positional variant is `Position(Position)` in core, pages,
  and keynote, but `Index(usize)` in numbers and `At`/`Index` in pages
  footnotes; the facade still takes raw `usize`; the core `Selector` carries a
  vestigial `Id` variant that xlsx already pins to `Infallible`.
- Default policies disagree on the lenient/refuse axis: docx
  `document_mut()` hydrates an opened package while pptx
  `presentation_mut()` refuses with `UnsafeEdit`; a changed write on a signed
  package is refused in xlsx (`Error::Signed`), requires an explicit `unsign()`
  in docx, and silently drops the signature in pptx (auto-`unsign()` inside
  commit); ODF builders' `save()` is a plain `std::fs::write` while OOXML
  `save()` is an atomic sibling-temp replacement; iWork builders consult
  ambient randomness for document identity, breaking byte-determinism; the
  facade reports iWork readers as `ToleranceReport::Untracked` (never clean)
  although those readers are fail-closed and have no repairs to track.

This record is the terminology-and-defaults counterpart of the CRUD scenario
checklist (ADR 0008): it fixes one canonical name per concept and one default
per policy, so that API users neither re-learn vocabulary per format nor guess
which lenient/refuse default applies.

## Decision

### 1. Layering principle

Structural vocabulary is unified library-wide; domain nouns are retained per
document family. One concept has exactly one name across all crates; a renamed
API is removed, never retained as an alias (ADR 0001 already declares backward
compatibility is not a constraint, and ADR 0004 already removes rather than
aliases).

- Universal: lifecycle verbs, transaction type family, error variant
  taxonomy, limits profiles, selector contract, metadata entry.
- Per-family: the semantic projection type name — `Document` (word
  processing), `Workbook` (spreadsheets), `Presentation` (decks).

### 2. Two-layer owner model

Every format exposes the same two layers:

- `Package` — the immutable, cheap-to-share owner snapshot. It owns every
  state verb: `open`, `new`, `edit`, `apply`, `save` and their variants. It is
  the ADR 0003 snapshot: mutation stages an `Edit`, `commit()` publishes a new
  `Package`, and the source snapshot is never mutated.
- `Document` / `Workbook` / `Presentation` — the family-named, owned,
  read-only semantic projection, obtained from the package
  (`Package::document()`, `workbook()`, `presentation()`). Owned projections
  are the primary form so `File` enum values and Python handles can hold them;
  borrowed views may exist as secondary `*_ref()` forms.

Consequences per family: `keynote::Document` is renamed `Presentation` and
`numbers::Document` is renamed `Workbook` (`pages::Document` already matches);
`xlsx::Workbook` is demoted to the read-only projection and its `edit`/`save`
verbs move to `xlsx::Package`; `odt::Document` and `odp::Presentation` split
into `Package` plus projection; `ods::Spreadsheet` is renamed `Workbook` and
demoted to the projection beside the existing `ods::Package`. The facade-level
`litchi::iwork::Document` unified reader is not renamed here; it is replaced
by the UNIFIED_OPS `File` enum in milestone M0a.

### 3. Transaction vocabulary

The iWork focused form is canonical:

- Scoped feature family: `module::{Edit, Patch, Commit, Diagnostics, Error,
  LimitKind}` with package entry points `Package::{x, edit_x, apply_x}`.
- The one primary document transaction uses the bare `Package::edit()`.
- `Edit::commit() -> Result<Commit>`; `Commit` always carries the triad
  `package()`, `patch()`, `diagnostics()`, plus `into_package()` and
  `into_parts() -> (Package, Patch)`. Every commit reports `Diagnostics`
  (xlsx and pptx gains it).
- Patch application is `Package::apply(&patch)` for the primary transaction
  and `apply_x(&patch)` for scoped families.
- No-op introspection is `commit.diagnostics().changed()` and
  `patch.is_noop()`; `patch.inverse()` is universal.

Renames and removals: pptx `Transaction` becomes `Edit`,
`opened_presentation_transaction()` becomes `edit()`, and `rollback()` is
deleted (dropping an owned `Edit` is the rollback); docx `edit_document()`
becomes `edit()`, and `publish_document_commit()` /
`apply_document_patch()` converge to `into_package()` / `apply()`; xlsx
`Commit::{workbook, into_workbook}` become `package`/`into_package` and
`Workbook::{edit, apply}` move to `Package`; xlsb `WorkbookEdit` /
`WorkbookCommit` / `WorkbookPatch` move to module-local short names;
`is_changed()`/`is_empty()` give way to the canonical introspection above.

### 4. Open, create, and save verbs

- Open: `open(path)`; `from_bytes(&[u8])`; `from_shared_bytes(Arc<[u8]>)`
  only where zero-copy sharing is supported; `from_reader(R)`; the configured
  entry is `open_with_options(path, OpenOptions)` /
  `from_bytes_with_options(bytes, OpenOptions)`. `from_slice`, `from_vec`,
  and ad-hoc `open_with` are deleted.
- Create: `Package::new()` produces the minimal valid artifact in every
  format; option-bearing authoring uses one `Builder` per format living in a
  focused module; `Builder::build()` returns `Package`, never raw bytes, and
  the builder has no `save`/`to_bytes` of its own. The xlsx `create()`
  duplicate is deleted; xlsb `WorkbookWriter` becomes `xlsb::Builder`; doc
  `Writer` becomes `doc::Builder`; the litchi-iwa `*DocumentBuilder` types
  drop their prefixes per the focused-module rule.
- Save: `save(path)` is unconditionally an atomic sibling-temp replacement in
  every format — the ODF package writer gains the sibling-temp + rename path
  that OOXML already has, and writing a partial output never touches the
  destination. `save_with_durability(path, Durability)` exposes the explicit
  downgrade; `to_bytes()` and `write_to(impl Write)` cover memory and stream
  output; `to_stream` and `write_to_stream` are deleted.
- Security verbs: `open_with_password(path, password)`,
  `save_encrypted(path, password, Mode)`, `save_reencrypted(path, password)`,
  `save_plain(path)` / `to_plain_bytes()`, and `encryption() -> Option<Mode>`.

### 5. Limits and configuration

- Each format exposes `SourceLimits` (physical ingress: bytes, members,
  depth, events) and, where a semantic layer exists, `SemanticLimits`
  (sheets, tables, cells, text bytes), plus one `OpenOptions` envelope per
  format with infallible consuming setters for pre-validated values.
- Renames: OPC `ReadLimits` becomes `SourceLimits` (docx/pptx/xlsx/xlsb
  re-exports follow); odt `SourcePackageLimits`, doc `package::Limits`, and
  iWork `DocumentSourceLimits` become `SourceLimits`; numbers
  `DocumentLimits` and keynote `DocumentSemanticLimits` become
  `SemanticLimits`; doc's `OpenOptions`/`PackageOpenOptions` pair merges into
  one `OpenOptions`; numbers `ReadOptions` and the facade `iwork::Options`
  take the `OpenOptions` name.
- Construction is a consuming builder with fallible `max_*` setters and a
  fallible `build()` (the OPC idiom); the iWork positional `new(...)`
  constructors are deleted.
- Errors: an invalid limit at construction is `Error::InvalidLimit { kind,
  value, maximum }`; an exceeded limit at ingress is `Error::Limit { kind,
  observed, maximum }`; kinds are `SourceLimitKind` / `SemanticLimitKind`
  enums, folded into the owning module's error type.
- Dimension names are nouns registered in the glossary (`input_bytes`,
  `members`, `entry_bytes`, `aggregate_bytes`, `depth`, `events`, …) with
  `max_<noun>` setters; format-specific dimensions may be added, but an
  existing dimension may not acquire a second name.
- Policy: caller-selected limits are mandatory on every public ingress path
  that accepts untrusted data (CRUD checklist row "Caller-selected limits").
  ODF and legacy gaps (ppt, xls) are filled; shared dimensions share default
  values across formats (for example 512 MiB input, XML depth 256).

### 6. Errors

- Structure: each focused module owns its non-exhaustive `Error`; the crate
  root keeps only ingress/egress error types. Boundary mapping uses typed
  variants (`#[from]` or explicit variants); string flattening of leaf errors
  (the `map_ooxml_error` pattern) is retired. The OOXML god-enums split per
  module, staged with the rollout waves. This amends ADR 0004's "each crate
  owns a typed error" clause, which its own iWork amendments had already
  superseded in practice.
- Shared variant taxonomy with fixed payload shapes:
  - `UnexpectedFormat { detected: FileFormat }` — recognized but wrong
    family; replaces `NotPages`/`NotKeynote`/`NotNumbers` and string
    messages. `FileFormat` is structural metadata, not authored content.
  - `NotOfficeFile` — unrecognized input.
  - `SourceChanged { expected, observed }` — the source file changed between
    open and use (TOCTOU).
  - `PatchConflict { .. }` — a patch's base does not match; absorbs
    `StaleSource`. The two names denote two distinct concepts and are not
    interchangeable.
  - `Limit { kind, observed, maximum }` and `Allocation { .. }` — per §5;
    `CollectionAllocation` folds into `Allocation`.
  - Expected failures are always typed; `Other(String)` narrows to genuinely
    internal errors, and the string-carrying legacy variants
    (`ParseError(String)`, `InvalidFormat(String)`) shrink per wave.
- Facade: `sheet::Workbook` returns typed `litchi_core::Error` instead of
  `Box<dyn Error + Send + Sync>`, and facade wrappers preserve leaf errors as
  typed sources. This is done now and subsumed by the UNIFIED_OPS error enum
  at M0a.
- Content redaction is a library-wide rule: errors, diagnostics, and `Debug`
  output never carry authored content (user text, entry names, sheet names);
  structural metadata (part paths, limit figures, positions, `FileFormat`)
  is allowed.

### 7. Metadata

- One `Metadata` type, merging `litchi_core::Metadata` and
  `ooxml_common::Props`; the `Props` type is deleted. Field reconciliation:
  `creator` becomes `author` (matching python-docx/python-pptx and user
  intuition; the `dc:creator` wire spelling stays at the codec boundary);
  `keywords` is the typed `Keywords`; timestamps are real time types, with
  W3c lexical forms confined to the codec. OOXML-only fields (`category`,
  `content_status`, `version`, `last_printed`) join `Metadata`; fields a
  format cannot store fail with a typed error or a recorded diagnostic, never
  silent loss.
- Entries: `Package::metadata()` is the universal read; docx's `props()` is
  deleted. Custom properties keep the unified `custom_props()` entry;
  application properties (`app.xml`) are an OOXML-specific `app_props()`
  layer; the physical core.xml part no longer has a separate typed channel —
  exotic needs go through the OPC raw layer.
- Mutation is transactional: `Package::{metadata, edit_metadata,
  apply_metadata}` plus `metadata::{Edit, Patch, Commit, Diagnostics, Error,
  LimitKind}`. `put_props()`, `props_mut()`, and `clear_props()` are deleted.
  ODF `Builder::set_metadata()` stays: it is fresh authoring, not opened
  mutation.
- Coverage: pptx and xlsx gain typed `metadata()`/`edit_metadata()`; iWork
  metadata write is a typed unsupported capability until its native
  acceptance gates pass.

### 8. Selectors

- The catalog selector is `Selector<'a> = Name(Cow<'a, str>) |
  Position(Position)`. The `Id` variant is deleted from the ordinary
  vocabulary; native-ID selection lives only in raw escape modules. Numbers'
  `Index(usize)` and the pages footnote `At`/`Index` forms converge on
  `Position`.
- `Position` is the only word for positional selection: a checked zero-based
  value type. Raw `usize` disappears from public signatures, including the
  facade's `slide(usize)` and `paragraph(usize)`.
- Methods accept `impl Into<Selector>` and return `Result<Option<_>>`;
  `From<&str>`/`From<usize>` keep call sites terse. Missing is `Ok(None)`;
  ambiguous and invalid are typed errors; no selector path panics or
  implements `Index`.
- Coordinates are not selectors: grid addressing uses the separate `At` /
  `Area` vocabulary (with A1 parsing at the edge); invalid syntax and
  out-of-grid coordinates are typed errors, never `None`.
- Name identity follows the producing application's semantics — xlsx sheet
  names are Unicode-caseless, pages section names are case-sensitive, docx
  glossary names use NFD/case-fold — and each format's policy is registered
  in the glossary. Fresh authoring without a native semantic defaults to
  canonical Unicode caseless identity.

### 9. Default policies

- Single write path: every mutation of any `Package` — fresh or opened —
  goes through `edit()`/`edit_x()` and `commit()`. The mutable writers
  (`document_mut()`, `presentation_mut()`) are deleted and their authoring
  verbs move into the corresponding `Edit` types. With no mutable writer,
  the `UnsafeEdit` refusal entry disappears by construction.
- Signatures: a changed write on a signed package refuses with
  `Error::Signed`; an exact no-op preserves signatures byte-identically;
  signature state changes only through the explicit verbs `sign()`,
  `resign()`, `unsign()`. pptx's automatic `unsign()` inside commit is
  deleted, and policy-parameter entry points (for example
  `PlaceholderSignaturePolicy`) converge on the verbs.
- Encryption: a plain `save()`/`to_bytes()` on an encrypted-source package
  refuses with `Error::EncryptionPolicy`; explicit downgrade is
  `save_plain()`/`to_plain_bytes()`; `save_reencrypted()` retains the source
  profile rather than silently dropping to defaults; unsupported encryption
  profiles are typed `Error::Unsupported`. This rule is registered now for
  the ODF and iWork write paths when they land.
- Determinism: builders and writers never consult the ambient clock,
  randomness, locale, or network. iWork document identity becomes an
  explicit required builder parameter: callers pick
  `Identity::generate()` (native-like fresh identity) or fixed bytes
  (deterministic pipelines and goldens); `build()` without an identity fails
  with a typed error. The OOXML/ODF fresh-creation paths are audited for
  hidden clock use (for example core.xml timestamps) and any found becomes
  explicit input.
- Tolerance reports: `ToleranceReport::Untracked(FileFormat)` is redefined
  as `Strict(FileFormat)` — a fail-closed reader that performs no repairs
  and therefore reports `is_clean() == true`. The rename is gated on an audit
  of the iWork read paths: any silent normalization found must become a
  recorded repair (a tracked report) or a typed error before the format may
  carry the `Strict` label.

### 10. Governance and rollout

- This ADR records the rationale; `docs/API_GLOSSARY.md` is the living
  per-term lookup table (canonical name, banned names, per-family status,
  wave). On acceptance of this record, ADR 0004 gains an amendment pointing
  its conflicting clauses here.
- Enforcement: a new `tools/check_api_glossary.py` generalizes the iWork
  rustdoc public-API gate pattern workspace-wide, with (a) a banned-name
  lint over public item names and (b) a shape check that any transaction
  family module exposes the complete `{Edit, Patch, Commit, Diagnostics,
  Error, LimitKind}` set. Each wave registers its banned names, making "is
  the rename done" a CI-decidable fact.
- Waves: W0 governance (this record, the glossary document, the checker
  skeleton); W1 iWork renames; W2 OOXML in docx → pptx → xlsx order; W3 ODF
  (package split, atomic save, limits coverage); W4 legacy (doc/ppt/xls).
  Facade-surface cleanup folds into the UNIFIED_OPS M0a migration so the
  facade breaks once, not twice; leaf-crate unification precedes M0a so the
  `File` enum maps onto already-unified names.
- Per-wave gates: warning-denied build, tests, and Clippy; checker and
  glossary status flips in the same commit; docs, README, and examples
  rewritten in the same commit; behavior-changing waves (ODF atomic save,
  iWork identity parameter) add CRUD-checklist evidence.

### 11. Deferred

History/undo-redo unification, patch composition verbs (`join`,
`three_way`), durable patch formats beyond the existing pptx/xlsx families,
the `litchi::iwork` facade reader, and the `litchi-ops` unified layer itself
remain on their own tracks (UNIFIED_OPS design, ADR 0003 deferrals).

## Consequences

- API users learn one verb set and one policy set; per-format learning is
  confined to domain nouns and format-specific capabilities.
- The iWork surface changes least in shape but takes the visible
  `Document`→`Workbook`/`Presentation` renames; the xlsx surface takes the
  deepest reshape (`Workbook` demoted to projection); docx/pptx lose their
  mutable writers, a broad but mechanical migration of authoring verbs into
  `Edit` types.
- Every rename is a hard removal; the 0.x series carries no deprecation
  window (ADR 0001). The checker prevents silent stragglers.
- Security-sensitive defaults move to the strict side (signatures refuse,
  encrypted sources refuse plain output, opened packages never hydrate into
  lossy writers); no convenience path silently weakens them.
- `docs/API_GLOSSARY.md` must be updated in the same commit as any public
  API naming change; drift between the table and the code is a CI failure
  once the checker lands.

## Verification

- W0: the glossary document and checker skeleton exist; the checker runs in
  CI over the current tree with an empty banned list.
- Each wave: `cargo build`/`test`/`clippy -D warnings` across the touched
  feature combinations; the checker with that wave's banned names passes;
  glossary rows flip in the same commit; examples and docs compile and run.
- Behavior changes carry evidence per the CRUD scenario checklist: ODF
  atomic save (failed-write leaves destination unchanged), iWork explicit
  identity (two builds with fixed identity are byte-identical; missing
  identity is a typed error), signature refusal (changed write on a signed
  package is `Error::Signed`; exact no-op preserves the signature), and
  encryption refusal (plain output of an encrypted source is
  `Error::EncryptionPolicy`).
