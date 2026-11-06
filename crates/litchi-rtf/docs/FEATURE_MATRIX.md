# Rich Text Format feature matrix

This matrix describes `litchi-rtf` against RTF 1.9.1 and [MS-OXRTFCP].
It records semantic APIs separately from exact source preservation. It does
not claim rendering fidelity, complete Word compatibility, or support for
every control word in a feature family.

## Reading this matrix

| Mark | Meaning |
|------|---------|
| ✅ | Public support within the stated scope |
| 🟡 | Partial, bounded, inert, or advanced retained-model support |
| ❌ | No typed semantic API for this direction |
| N/A | Direction does not apply |

`Package` is the immutable lifecycle and publication owner.
`Package::document()` returns a cheap owned read-only semantic projection, and
`Package::edit()` publishes source-checked changes. `raw::Document` and the
native value vocabulary expose the advanced retained model; a retained-model
writer is not an ordinary transactional CRUD API. Rows explicitly name that
distinction where relevant. Untouched `Package` serialization retains the original bytes, including
unknown syntax. Canonical retained-model serialization may change lexical
spelling and refuses opaque structures whose ownership it cannot preserve.
An unsupported semantic row therefore does not imply loss on an untouched save.

## Input, output, and transactions

| Feature | Read | Write | Scope and evidence |
|---------|------|-------|--------------------|
| RTF groups, control words, escaped text, binary payloads | ✅ | ✅ | Bounded lexer/parser and `write::Writer`; source, token, nesting, binary, and opaque-data limits. `tests/parse_limits.rs`, `group_nesting_depth.rs`, `robustness.rs`. |
| Unicode and legacy code pages | ✅ | 🟡 | Unicode fallback counts and font-selected code pages are decoded, including CP932 in secondary stories. Untouched source bytes survive. Changed byte-oriented sources can be refused; canonical output is not an exact lexical rewrite. `story_font_codepages.rs`, `transport_and_producer_contracts.rs`. |
| Immutable package and read-only projection views | ✅ | N/A | `Package::document`, `Package::body`, `Package::fonts`, `Package::colors`, and borrowed paragraph/run traversal. A projection is an owned read-only handle over the retained package, not bounded-memory streaming input. `facade.rs`, `sequential_text.rs`. |
| Streaming fresh creation | N/A | 🟡 | `streaming::Writer` writes bounded paragraph/run events to a sequential sink; `write::Writer` offers text/formatting and selected note, hyperlink, and revision helpers. No complete ordinary fresh table/list/shape/math builder. |
| Existing-package transactions | ✅ | 🟡 | Disjoint body spans, paragraph layout/lifecycle including positioned-frame create/update/remove, bold/italic ranges, table-cell text, header/footer text, annotations, notes, root shape text, and bounded picture payload replacement/removal. Preconditions and supported owner closures are operation-specific; frame edits use the ordinary source-bound layout transaction, preserving checked opaque body spans and refusing unretained or dependent syntax. Table-cell, header/footer-paragraph, note and comment text publish by splicing the replacement into parser-recorded source provenance (`destination_splice.rs`), so tables keep their body position and untouched bytes, unknown destinations included, are copied verbatim; ranges with state-changing or destination controls are refused, and the canonical writer remains only for other destinations in table-free documents. `Package::edit()` and the transaction suites cover publication. `transactions.rs`, `paragraph_layout_transactions.rs`, `paragraph_frames.rs`, `picture_payload_crud.rs`, `picture_removal_crud.rs`. |
| Reversible patches, composition, and history | ✅ | ✅ | Exact-source guards, durable replay/inverse, conservative disjoint composition, explicit conflict resolution, and bounded history for supported edits. `durable_composition.rs`, `transactions.rs`. |
| Cross-document transfer | ✅ | 🟡 | Checked ordinary-root passive field, nested-table, style, list/picture-bullet, shape, and embedded-object closures. Active links, unknown destination ownership, and unresolved resource collisions are refused. `ordinary_root_transfer.rs`. |
| Tail append | ✅ | 🟡 | `tail_append` supports checked logical/source append with explicit limits; this does not authorize arbitrary structure rewrites. `logical_tail_append.rs`, `source_tail_append.rs`. |
| Unknown control words and starred destinations | 🟡 | 🟡 | `opaque::Node` keeps bytes and owner anchors. Exact snapshot output and eligible body splices preserve them; ordinary paragraph-layout edits reinsert body-anchored nodes at their checked source spans, while other canonical edits fail closed when ownership cannot be proved. `opaque_preservation.rs`, `paragraph_layout_transactions.rs`. |
| Validation and redaction | 🟡 | 🟡 | Structured validation/security reports and bounded explicit redaction paths. No general repair, sanitization certification, or unknown-syntax interpretation. `validation_report.rs`, `external_reference_redact.rs`. |
| Compressed RTF transport, [MS-OXRTFCP] | ✅ | ✅ | `transport::{compress,decompress}` handles LZFu/MELA transport, checksum/framing checks, and requires `CompressionLimits`/`DecompressionLimits` profiles. Both defaults select a 256 MiB output ceiling; `new(max_output_bytes)` lets callers choose another finite bound, which is checked before fallible output growth and final frame allocation. Exact retained compressed snapshots are writable; ordinary changed compressed snapshots are refused. `transport_and_producer_contracts.rs`, `src/codec/compressed.rs`. |

## Document and content vocabulary

| Feature | Read | Write | Scope and evidence |
|---------|------|-------|--------------------|
| Character formatting and paragraph properties | ✅ | 🟡 | Typed fonts, colors, decorations, borders/shading, tabs, spacing, indentation, direction, and break controls. Selected ordinary edit operations; broader retained-model write-back. `character_decorations.rs`, `paragraph_tabs.rs`, `paragraph_spacing_policy.rs`, `paragraph_layout_transactions.rs`. |
| Paragraph, line, page, column, and soft breaks | ✅ | ✅ | Distinct retained structural boundaries; no pagination engine. `body_story_event_order.rs`, `page_breaks.rs`, `column_breaks.rs`, `soft_breaks.rs`. |
| Font/color tables and embedded fonts | ✅ | 🟡 | Borrowed catalogs, sparse IDs, font metadata/theme references, and retained embedded programs. No font rendering or unrestricted ordinary font CRUD. `font_metadata.rs`, `embedded_fonts.rs`, `font_theme.rs`. |
| Stylesheets, Quick Styles, and personal/compose/reply styles | ✅ | 🟡 | Paragraph/character/section/table styles, links, inheritance checks, `\sqformat`, `\spersonal`, `\scompose`, and `\sreply` are modeled and emitted by retained writers. `stylesheets.rs`, `src/model/document/tests.rs`. |
| Table styles and conditional formatting | ✅ | 🟡 | Retained style references, conditional regions, and banding. Native Word definitions sharing a table-style ID retain their distinct conditional scopes through reads, stylesheet writes and style transfers; no table layout engine. `table_style_references.rs`, `table_style_conditional.rs`, `table_autoformat_banding.rs`, `microsoft_office_native.rs`. |
| Latent styles and style/format restrictions | ✅ | 🟡 | `DocumentStyleRestrictions` and retained policy values preserve restrictions as metadata; general policy mutation is not ordinary CRUD. `latent_styles.rs`, `document_style_restrictions.rs`, `style_list_filter.rs`. |
| Tables, nested tables, merges, and floating tables | ✅ | 🟡 | Retained rows/cells, nesting, merges, borders, widths, positioning, and metadata. Bounded exact-source cell-text editing and checked table transfer; no general fresh ordinary table builder or table/row/cell structure editor. `destination_splice.rs`. Floating-table support does not cover paragraph frames. `nested_tables.rs`, `table_merge.rs`, `floating_table_positioning.rs`, `table_geometry.rs`. |
| Lists and picture bullets | ✅ | 🟡 | Levels, overrides, legacy numbering, markers, and picture resources are modeled and can be retained or transferred within checked closures. No general ordinary list authoring API. `lists.rs`, `list_pictures.rs`, `legacy_numbering.rs`, `generated_list_markers.rs`. |
| Sections and header/footer stories | ✅ | 🟡 | Page/column/note settings and story ownership are typed. Bounded header/footer text transactions; retained write-back of other settings. `section_columns.rs`, `section_note_options.rs`, `drawing_story_ownership.rs`. |
| Footnotes and endnotes | ✅ | 🟡 | Typed note stories/options/separators and bounded exact-source story edits. `\ftnalt` marks an endnote, and the canonical writer emits endnotes as `\footnote\ftnalt`. `note_options.rs`, `note_separators.rs`, `transactions.rs`, `story_references.rs`. |
| Comments/annotations | ✅ | 🟡 | Retained annotation metadata and bounded comment-body transactions. No collaboration engine. `annotations.rs`, `transactions.rs`. |
| Bookmarks and navigation entries | ✅ | 🟡 | Plain bookmarks, TOC/index navigation metadata, and cached content are retained. No index generation or field calculation. `navigation_entries.rs`, `table_story_navigation_revisions.rs`. |
| Tracked revisions | ✅ | 🟡 | Insertion/deletion and structural revision metadata plus selected writer helpers; no general ordinary accept/reject or merge engine. `revisions.rs`, `structural_revisions.rs`. |
| Fields and hyperlinks | ✅ | 🟡 | Instruction/cache/status models and inert typed references. Starred body bookmark destinations inside native field instructions retain their spans without polluting the instruction text; non-body bookmarks inside instructions stay opaque. Retained writing and checked passive field transfer; never refresh or evaluate. `field_status.rs`, `typed_hyperlinks_and_references.rs`. |
| Form fields and editable/protected regions | ✅ | 🟡 | Inert control properties, protected users/ranges, and editable regions are modeled and retained. Native Word text defaults/help/max length, checkbox default/size, drop-down entries and undefined selections survive canonical writing with their result ranges and bookmarks. Form-owner modifier records retain native `flddirty`, `fldedit`, `fldlock` and `fldpriv` independently of the generic-field projection; adjacent empty forms keep distinct flags and newly authored forms use defaults. Default values are not substituted for absent results; controls never activate. `form_field_properties.rs`, `protection_ranges.rs`, `editable_regions.rs`, `microsoft_office_native.rs`. |
| Math and equations | ✅ | 🟡 | Retained math trees/properties and writer support; no ordinary fresh math builder, evaluation, or layout. Native Word `mmath` envelopes, inline/display trees, character-format wrappers and both script-child orders are covered by `microsoft_office_native.rs`. Multiple equations in one display paragraph return typed `UnsupportedContent`; local run formatting and picture fallbacks remain source-owned. `math_zones.rs`, `math_properties.rs`. |
| Pictures, drawings, shape groups, and text boxes | ✅ | 🟡 | Typed retained metadata and payloads, bounded picture edits, root-shape text transactions, and checked transfer. No visual renderer or arbitrary shape graph editor. `picture_properties.rs`, `shape_groups.rs`, `shape_text_frames.rs`, `legacy_drawings.rs`. |
| OLE/embedded objects and Publisher objects | 🟡 | 🟡 | Inert object class/data/result metadata includes `ObjectKind::Publisher` (Macintosh Edition Manager). Nested result formatting groups retain scoped text decoding and preferred/fallback picture wrappers in `result_picture_wrappers`, distinct from body picture records. Canonical writing and passive transfer retain those wrappers and payloads; objects never activate or convert. `objects.rs`, `src/drawing/object.rs`, `microsoft_office_native_visuals.rs`. |
| Document information and producer metadata | ✅ | 🟡 | Retained information, user properties, variables, generator, origin, file tables, themes, and XML namespace/data-store metadata. Empty `datastore` destinations emitted by native Word remain present through reads, exact-source edits and canonical writing. `document_info.rs`, `user_properties.rs`, `document_variables.rs`, `custom_xml.rs`, `data_store.rs`, `microsoft_office_native.rs`. |
| Mail merge, external references, and XSL transforms | 🟡 | 🟡 | Metadata and cached values stay inert; native Word empty file tables, main/input type, destination, current record and error-mode metadata survive read/write, including unstarred query/source text groups. Retained write-back and explicit bounded external-reference redaction. No source access, transformation execution, or merge execution. `mail_merge.rs`, `external_references.rs`, `xsl_transform.rs`, `microsoft_office_native.rs`. |
| Document layout/compatibility/view/print policies | 🟡 | 🟡 | Typed retained policy families have codec support, not host behavior or ordinary general policy CRUD. `document_compatibility_policy.rs`, `document_view.rs`, `document_output_settings.rs`. |
| Bidirectional, associated-font, and East Asian formatting | 🟡 | 🟡 | Modern direction, associated fonts, custom kinsoku, grid, and expansion metadata; CP932 inference from selected font charset is supported. Word 6J legacy controls below remain separate. `bidirectional.rs`, `associated_character_formatting.rs`, `kinsoku.rs`, `story_font_codepages.rs`. |
| Legacy document protection password | 🟡 | 🟡 | `\password` hex and protection flags are retained as inert metadata; values are not authenticated. Supported transaction paths enforce their own protection preconditions. `document_protection.rs`, `paragraph_batch_and_protection.rs`. |
| Write-reservation metadata | 🟡 | 🟡 | `\writereservation` and `\writereservhash` have bounded opaque payload models and retained write-back. These are distinct from `\passwordhash`. `write_reservations.rs`, `src/metadata/write_reservation.rs`. |

## Explicit semantic gaps and exclusions

| Feature | Read | Write | RTF 1.9.1 scope and current boundary |
|---------|------|-------|-------------------------------------|
| Positioned paragraph objects and frames | ✅ | 🟡 | pp. 91–93: typed `ParagraphFrame` metadata covers size, anchor references/positions, text distances, wrapping, overlay, no-overlap, locking, and text-flow controls. Retained-model writing plus ordinary source-bound paragraph-layout create/update/remove transactions are supported; checked source spans preserve body opaque controls and untouched metadata, while unretained/dependent syntax is refused atomically. No layout engine is provided. Drop caps, floating tables, and drawing text boxes remain separate. `paragraph_frames.rs`, `paragraph_layout_transactions.rs`. |
| SmartTag/factoid data | ✅ | 🟡 | p. 153: main-body starred `\xmlopen`/`\xmlclose` with required `\xmlnsN`, `\factoidname`, and bounded `\xmlattr` groups whose required namespace control is `\xmlattrnsN` are exposed as inert, source-bound `SmartTag` ranges. Illegal parameters, namespace/control aliases, value-before-name attributes, and unclosed markers are rejected. Canonical writing and `Package::edit().set_smart_tag` preserve the covered body range; unknown factoid actions, namespace resolution, and execution remain outside the API. `smart_tags_move_bookmarks.rs`. |
| Move bookmarks | ✅ | 🟡 | pp. 146–148: main-body-only starred `\mvfmf`/`\mvfml` and `\mvtof`/`\mvtol` start/end ranges retain alphanumeric tags and the six-byte author/DTTM payload as inert `MoveBookmark` metadata. Equal tags with opposite kinds identify the two move locations, while the bounded model does not expose a cross-location pair index or apply the specification's deleted/inserted fallback. Duplicate same-location starts and numeric move-control parameters are rejected; unmatched markers remain semantically ignored, and changed source-bound or canonical retained-model publication refuses their source because the recognized marker cannot be emitted safely. Move execution is outside the API. `smart_tags_move_bookmarks.rs`. |
| Modern protection password hash | ✅ | 🟡 | pp. 40–41: starred root-header `\passwordhash` SDATA is retained as an inert, bounded `PasswordHash` record with typed views over the observed fixed header and logical salt/hash ranges, plus exact unknown/trailing-byte preservation. `PasswordHash::from_parts` accepts caller-supplied header and tail data; the convenience authoring constructor makes no Word-verifier compatibility claim. The record is never authenticated or executed. `\password` and `\writereservhash` remain distinct. `document_protection.rs`. |
| Word 6J-era East Asian controls | ❌ | ❌ | pp. 204–206: `\jsksu`, `\jsku`, `\horzvert`, `\gcw`, `\twoinone`, `\nosectexpand`, `\sectexpand`, and `\jclisttab` lack typed semantics. |
| Rendering, pagination, shaping, and field/formula calculation | ❌ | ❌ | Deliberate library boundary; cached values and layout metadata are not computed results. |
| Macro/control/object execution and external refresh | ❌ | ❌ | Deliberate boundary. Retaining links or embedded bytes never causes network access or execution. |

## Verification boundary

Evidence paths above are relative to `crates/litchi-rtf`; bare test filenames
refer to `tests/`. They identify executable coverage, not a claim that every
test or native Office gate was run when this matrix was added. The transport
suite contains specification vectors and checked-in native-resave artifacts;
those fixtures do not prove compatibility with every producer/version.
No performance improvement is claimed by this documentation change.

The native Word embedded-font derivative has 91 sparse font definitions and
ten empty `fontemb` declarations. Its `ftruetype` spelling is accepted as an
import alias; canonical writing emits `fttruetype`. Exact names, theme roles,
absent payload/file fields and font metadata survive writer and body-run edits.
Empty declarations are not font-extraction evidence. Source-run matching uses
the parser's font-codepage/charset precedence and scoped document encoding,
including multibyte escapes; physical CR/LF outside visible text stays intact.
The public facade regression checks the native Aptos property and text edit
without losing unrelated symbol text or font declarations. Registration,
substitution, licensing and rendering remain outside this evidence.

`tests/microsoft_office_native_visuals.rs` qualifies the native Word worksheet
and chart OLE descriptors, complete objdata hashes, bounded OLE1 headers and
embedded CFB bytes. Each object's preferred/fallback previews survive model
writing and transfer into an ordinary generated receiver. Unknown result
destinations retain their structural scope and block lossy canonical output.
The native vector specimen's PNG bytes remain distinct from its declared
`pngblip`/`wmetafile8` transport types. Its field-scoped `defshp` empty-picture
metadata remains unmodeled: canonical rewriting refuses before output.

Exact snapshot save/reopen retains both visual specimens byte for byte.
Ordinary-run edits still refuse these dependent picture/object/field owners,
and ordinary-root transfer refuses a native receiver with opaque destinations;
both boundaries have typed refusal and unchanged-source assertions. No native
visual edit, empty-default-shape authoring, rendering or object activation is
inferred from the successful snapshot and transfer cases.

Native layout regressions qualify document page-default inheritance, section
geometry, grouped section boundaries, explicit page breaks, running header/footer
stories, note bodies/separators and selected passive policies. Document page
controls and their section overrides have separate tokens; `sectd` restores the
document geometry, and explicit section overrides remain distinct. `sectdefaultcl`
sets the default document grid without resetting columns or line pitch.
`SectionDocumentGridType::Default` now serializes as `sectdefaultcl`; the separate
`SnapToGrid` variant uses the specification's literal `sectspecifygenN`, with the
older `sectspecifygen` spelling accepted on input.

Native header, footnote/endnote and section-value edits preserve unrelated
source bytes and invert after reopen. Section edits follow retained definition
identity across ordinary formatting groups and inherited sections. Balanced
`fcs` character-script groups are admitted by destination text splicing; controls
that change state outside a balanced selected group remain subject to the
existing exact-source/canonical-publication boundaries. These tests do not claim
pagination, PAGE/note-number evaluation, note-area columns, general policy CRUD
or fresh Word acceptance of edited files.

`microsoft_office_native_workflows.rs` qualifies bounded logical/source-backed
plain append on native Word sources, durable replay/inverse, finite publication
windows and a native-to-native table-cell text transfer. Protected, dropped-syntax
and external-dependency sources retain precise refusal tests. The separate
streaming case produces a plaintext derivative of native multilingual content;
it does not copy the original formatting or layout.

`microsoft_office_native_validation.rs` distinguishes parser-proven native counts
from unresolved opaque scope. Its strict redaction cases refuse, and best-effort
no-ops remain explicitly incomplete; no complete native external-content removal
is claimed. LZFu/MELA cases wrap native payload bytes in library-generated
transport, with byte-exact decompression, finite bounds and checksum checks.
The native corpus does not contain Office-produced compressed-RTF transport.
