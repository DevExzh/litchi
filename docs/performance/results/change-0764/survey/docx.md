# litchi-docx: quick-xml attribute-iterator call sites (survey for 0764)

Worktree: `/home/zhuhe/code/litchi-worktrees/0764-xml-attribute-dos-hardening` at `1d1044e3ac`; quick-xml 0.41.0.
Scope: every `.attributes()` / `.html_attributes()` in `crates/litchi-docx/src/**/*.rs`. The grep finds 175 occurrences
and no `.html_attributes()`. There are no `try_get_attribute` calls, and `::attributes` paths appear only as type paths.

Facts checked in the quick-xml 0.41.0 source:
- `BytesStart::attributes()` returns an iterator with `check_duplicates: true`. Past 32 keys it uses an identity-hashed prefilter.
  Every real duplicate then triggers a linear scan of the earlier keys. After an `Err(Duplicated)` the iterator moves on to the next attribute.
- `Attributes` implements only `next`, with no `size_hint` override, so `size_hint().0` is always 0. `.count()` walks every item.
- `NamespaceResolver::resolve_attribute` does a reverse linear scan over the in-scope bindings. By default NsReader caps
  declarations at 256 per element, but the total grows with depth. `NamespaceResolver::bindings()` runs `NamespaceBindingsIter::next`,
  which rescans all later bindings to detect overrides, so a full traversal costs O(B^2).

Legend:
- checks: `checked` means the default or `.with_checks(true)`; `(explicit)` marks the latter. `unchecked` means `.with_checks(false)`.
- mode: FAIL-FAST, TOLERANT, SHORT-CIRCUIT and COUNT, as defined in the task.
- consumption: NAMES/ALL/COUNT/OTHER. "guarded" means a later match for the same (expanded) name is refused.
  "OVERWRITE" means plain assignment, so the last match would win if duplicates were yielded.
- can-refuse: `yes` when the enclosing fn returns `Result` (or `Option<Result>`).
- Row paths are relative to the worktree root.

## Site table

| file:line | fn | checks | error mode | consumption | can-refuse | own-duplicate-check | other super-linear | reachability |
|---|---|---|---|---|---|---|---|---|
| `crates/litchi-docx/src/alt/codec.rs:318` | relationship | checked | FAIL-FAST | NAMES r:id (guarded) | yes | O(1) guard on expanded r:id | none | altChunk elements in document XML |
| `crates/litchi-docx/src/alt/codec.rs:358` | parse_on_off | checked | FAIL-FAST | NAMES w:val (guarded) | yes | O(1) guard | none | altChunk matchSrc |
| `crates/litchi-docx/src/bibliography/mod.rs:461` | xml_node | checked | FAIL-FAST | ALL non-xmlns attrs to Vec | yes | none (relies on quick-xml check) | none | bibliography customXml DOM parse |
| `crates/litchi-docx/src/bookmark.rs:85` | extract_from_document | checked | **TOLERANT** (`.flatten()`) | NAMES id,name by local name; OVERWRITE | yes | none | none | `Package::bookmarks()` over raw document.xml; open does no attribute validation |
| `crates/litchi-docx/src/chart/codec.rs:380` | element_info | checked (explicit) | FAIL-FAST | ALL expanded attrs to Vec | yes | `values.iter().any` per attr: **QUADRATIC** on n distinct attrs; only caps are MAX_ATTRIBUTES=750,000 per part and 16 MiB attr bytes | (the own check) | chart graph scans: `scan_document_xml` (document.xml), `scan_chart_xml` (chart parts) |
| `crates/litchi-docx/src/comment.rs:273` | parse_comment_metadata | checked | FAIL-FAST | NAMES id/author/initials/date by local name; plain assignment | yes | none | none | comments.xml, each comment start tag |
| `crates/litchi-docx/src/content_control.rs:1577` | extension_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) guard | none | content-control extraction (ext attrs) |
| `crates/litchi-docx/src/content_control.rs:1612` | exact_extension_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) guard | none | same |
| `crates/litchi-docx/src/content_control.rs:1836` | choice_is_supported | checked | FAIL-FAST | NAMES raw `Requires`; plain assignment | yes | none | none | content-control mc:Choice |
| `crates/litchi-docx/src/content_control.rs:1886` | extend_ignorable | checked | FAIL-FAST | OTHER (each mc:Ignorable split into prefixes; bounded by metadata/binding limits) | yes | none | none (IgnorableState is O(1)) | content-control MCE scope |
| `crates/litchi-docx/src/content_control.rs:1962` | require_ignorable_extensions | checked | FAIL-FAST | NAMES storeItemChecksum/formattingAllowed (bool guards) | yes | 2 bool flags, O(1) | none | content-control sdtPr children |
| `crates/litchi-docx/src/content_control/snapshot.rs:663` | lock_value | **unchecked** | FAIL-FAST (syntax Err only) | NAMES w:val (guarded) | yes | `value.is_some()` for expanded w:val only; other duplicate names are not detected | none | content-control snapshot (w:lock) |
| `crates/litchi-docx/src/content_control/snapshot.rs:731` | mce_directives | **unchecked** | FAIL-FAST (syntax Err only) | OTHER (every mc:Ignorable/ProcessContent attr processed; repeats merged) | yes | none for attributes. Ignorable namespaces deduped via `local_ignorable.iter().any` (Vec) | Vec `any` per Ignorable token: quadratic in distinct namespaces (tokens <= max_bindings=65,536). Plus an `ancestors x introduced_ignorable` scan per namespace | content-control snapshot MCE frame, every element |
| `crates/litchi-docx/src/content_control/snapshot.rs:914` | exact_attribute | checked | FAIL-FAST | COUNT cap 1,024 checked during iteration (Ok items only) + NAMES (first span kept, match count returned) | yes | none (count returned to caller) | `find_attr` rescans the whole tag lexically for each matching attr (<=1,024) | content-control snapshot (4 calls per element) |
| `crates/litchi-docx/src/document/transaction.rs:6156` | event_namespace_binding_count | **unchecked** | COUNT (tolerant `.filter_map(Result::ok)`; counts Ok only) | COUNT xmlns bindings | no (usize) | none | none (linear; skipped unless the raw attrs contain `xmlns`) | managed document.xml scan (`scan_document_with_context`) in transactions |
| `crates/litchi-docx/src/document/transaction.rs:6525` | oracle_event_namespace_binding_count | unchecked | COUNT (tolerant) | COUNT | no | none | none | **TEST** (`#[cfg(test)]` oracle fn) |
| `crates/litchi-docx/src/document/transaction.rs:7520` | validate_revision_element_attributes | checked | FAIL-FAST | COUNT cap MAX_REVISION_ATTRIBUTES=256 checked during iteration (Ok only) | yes (Refusal) | none | none | transaction revision elements |
| `crates/litchi-docx/src/document/transaction.rs:7669` | raw_revision_scope_attributes | checked (never iterated) | n/a: `size_hint().0` only | OTHER: always 0 (no `size_hint` impl), so `try_reserve_exact(0)` is a no-op | yes | n/a | none | same fn as next row |
| `crates/litchi-docx/src/document/transaction.rs:7678` | raw_revision_scope_attributes | checked | FAIL-FAST | COUNT cap 256 (Ok only) + NAMES (bool guards) + collected scope attrs | yes | 5 bools + `seen_scope_attributes.iter().any` over <=5 entries (O(1)) | `raw_tag_attributes` lexical scan once (linear) | transaction revision scope attributes |
| `crates/litchi-docx/src/document/transaction.rs:8771` | complex_field_marker | checked | **SHORT-CIRCUIT** (`.filter_map(Result::ok).find(local==fldCharType)`): quadratic when fldCharType is absent (the result is ComplexContent) or last | NAMES single | yes (Refusal) | none | caller `select_complex_field_result` reparses the paragraph from the start for each run (O(runs x bytes), runs <= MAX_OPERATIONS=4,096). Also `resolver().clone()` per event | paragraph-edit transactions that select complex-field results (run XML) |
| `crates/litchi-docx/src/drawing/codec.rs:326` | parse_anchor_id | checked | FAIL-FAST | NAMES *:anchorId (guarded) | yes | O(1) guard | none | `drawing::parse` (paragraph drawings) |
| `crates/litchi-docx/src/drawing/codec.rs:373` | inert_attribute | checked | **TOLERANT** (`.flatten()`; always scans everything, no early exit) | NAMES exact raw key; OVERWRITE | no (Option; callers return Result) | none | helper called up to 8 times per drawing (6 via `number_attribute`, plus name and descr), each a full scan | `drawing::parse` from Paragraph content (`paragraph/codec/content.rs:128`) |
| `crates/litchi-docx/src/drawing/codec.rs:387` | strict_attribute | checked | FAIL-FAST | NAMES raw key; plain assignment | yes | none | none | prstGeom@prst |
| `crates/litchi-docx/src/drawing/source.rs:2107` | collect_relationship_attributes | checked (explicit) | FAIL-FAST | OTHER (all r:* refs collected; capped by max refs) | yes | none | none | source-backed drawing index of main story |
| `crates/litchi-docx/src/drawing/source.rs:2525` | svg_relationship_dialect | checked (explicit) | FAIL-FAST | NAMES *:embed/*:link (`dialect.replace` guard) | yes | O(1) | NESTED: `relationship_dialect_for_prefix` rescans all attrs for each *:embed/*:link attr (O(k*n); n <= 256 because `validate_attributes` already ran on the same tag) | SVG blip fragment of an admitted SVG owner |
| `crates/litchi-docx/src/drawing/source.rs:2566` | relationship_dialect_for_prefix | checked (explicit) | FAIL-FAST (returns at the matching xmlns:prefix) | NAMES | yes | n/a | then `context.visit_visible` (all visible bindings) | called once per attribute from the 2525 loop |
| `crates/litchi-docx/src/drawing/source.rs:2606` | extend_namespace_context | checked (explicit) | FAIL-FAST | ALL xmlns declarations | yes | none | none | drawing index, elements with declarations |
| `crates/litchi-docx/src/drawing/source.rs:2635` | namespace_declaration_bytes | checked (explicit) | FAIL-FAST | OTHER (byte sum) | yes | none | none | drawing index, every element |
| `crates/litchi-docx/src/drawing/source.rs:2662` | validate_attributes | checked (explicit) | FAIL-FAST | COUNT cap MAX_ATTRIBUTES=256 during iteration (Ok only) + expanded names | yes | `expanded_names.iter().any` (Vec). Quadratic but n <= 256, so <= ~32k compares | none | drawing index, every element |
| `crates/litchi-docx/src/drawing/source.rs:2797` | extension_uri | checked (explicit) | FAIL-FAST | NAMES `uri` (`found` flag; an extra one is marked malformed) | yes | bool | none | a:ext |
| `crates/litchi-docx/src/drawing/source.rs:2846` | relationship_attribute | checked (explicit) | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | blip embed/link |
| `crates/litchi-docx/src/drawing/source.rs:2887` | unqualified_attribute | checked (explicit) | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | docPr id/name, a:ext uri |
| `crates/litchi-docx/src/drawing/validation.rs:44` | parse_word2010_anchor_id | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | legacy VML anchorId |
| `crates/litchi-docx/src/field/codec.rs:86` | extract_from_document | checked | FAIL-FAST | NAMES fldCharType/dirty/fldLock by local name; plain assignment | yes | none | none | field extraction over document XML |
| `crates/litchi-docx/src/field/codec.rs:1352` | PendingSimpleField::parse | checked | FAIL-FAST | NAMES instr/dirty/fldLock; plain assignment | yes | none | none | w:fldSimple |
| `crates/litchi-docx/src/font/codec.rs:148` | make_node | checked (explicit) | FAIL-FAST | ALL | yes | `HashSet<String>` (std SipHash), O(n), redundant with checks | per-element clone of the inherited scope HashMap | fontTable.xml DOM parse |
| `crates/litchi-docx/src/font/open_type/codec.rs:605` | parse_root_attributes | checked (explicit) | FAIL-FAST | NAMES val (guarded; unknown names refused) | yes | O(1) | none | w14 OpenType run properties |
| `crates/litchi-docx/src/font/open_type/codec.rs:674` | parse_style_set | checked (explicit) | FAIL-FAST | NAMES id/val (guarded; unknown names refused) | yes | O(1) | none | w14:styleSet |
| `crates/litchi-docx/src/font/package.rs:852` | directly_used_font_names | checked (explicit) | FAIL-FAST | NAMES rFonts ascii/hAnsi/eastAsia/cs into a HashSet | yes | none | none | font-usage validation over every WordprocessingML part |
| `crates/litchi-docx/src/footnote.rs:375` | parse_note_metadata | checked | FAIL-FAST | NAMES id/type; plain assignment | yes | none | none | footnote/endnote start tags |
| `crates/litchi-docx/src/glossary/codec/xml.rs:333` | make | checked (explicit) | FAIL-FAST | ALL; cap MAX_VALUES=4,096 checked before each item | yes | none | none (HashMap resolver) | glossary document DOM parse |
| `crates/litchi-docx/src/hyperlink.rs:190` | extract_from_xml | checked | FAIL-FAST | NAMES id/anchor/tooltip; plain assignment | yes | none | none | hyperlink extraction |
| `crates/litchi-docx/src/image.rs:308` | parse_inline_images | checked | **TOLERANT** (`.flatten()`) | NAMES cx/cy (raw key); OVERWRITE | yes | none | none | `Paragraph` images (`paragraph/codec/content.rs:98`), run contents (`run_contents.rs:282`) |
| `crates/litchi-docx/src/image.rs:327` | parse_inline_images | checked | **TOLERANT** (`.flatten()`) | NAMES name/descr; OVERWRITE | yes | none | none | same |
| `crates/litchi-docx/src/image.rs:349` | parse_inline_images | checked | **TOLERANT** (`.flatten()`) | NAMES `*:embed` suffix; OVERWRITE (distinct prefixes overwrite even with checks on) | yes | none | none | same |
| `crates/litchi-docx/src/ink/codec.rs:698` | choice_requires_kind | checked | FAIL-FAST (`attribute.ok()?` returns None) | NAMES Requires (a second one also returns None) | no (Option) | O(1) | none | ink story mc:Choice |
| `crates/litchi-docx/src/ink/codec.rs:835` | read_anchor | checked | FAIL-FAST | NAMES r:id (guarded) | yes | O(1) | none | ink contentPart anchor |
| `crates/litchi-docx/src/ink/codec.rs:1072` | graphic_data_kind | checked | FAIL-FAST (returns None) | NAMES uri (a second one returns None) | no (Option) | O(1) | none | ink graphicData |
| `crates/litchi-docx/src/ink/host.rs:984` | graphic_data_kind | checked | FAIL-FAST (returns None) | NAMES uri (a second one returns None) | no (Option) | O(1) | none | ink host graphicData |
| `crates/litchi-docx/src/ink/host.rs:1392` | validate_group_metadata | checked | FAIL-FAST | COUNT (non-xmlns count compared with the expected value after the loop) + NAMES | yes | none | none (preceded by `ink::xml::element`, 256 cap) | ink group metadata |
| `crates/litchi-docx/src/ink/host.rs:1458` | collect_attributes | checked | FAIL-FAST | ALL values tokenized into `HashSet<String>` | yes | HashSet (std) | none (preceded by `element`, 256 cap) | ink host relationship references |
| `crates/litchi-docx/src/ink/package.rs:1109` | emma_mode_is_ink | checked | FAIL-FAST | NAMES emma:mode (guarded) | yes | O(1) | none | InkML part |
| `crates/litchi-docx/src/ink/package.rs:1135` | validate_profile_attributes | checked | FAIL-FAST | COUNT cap MAX_PROFILE_ATTRIBUTES=256 (Ok only) | yes | none | none | InkML profile |
| `crates/litchi-docx/src/ink/package.rs:1209` | xml_id | checked | FAIL-FAST | NAMES xml:id (guarded) | yes | O(1) | none | InkML definitions |
| `crates/litchi-docx/src/ink/package.rs:1229` | profile_attr | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | InkML trace attrs |
| `crates/litchi-docx/src/ink/placement.rs:1215` | choice_requires_supported | checked | **SHORT-CIRCUIT** (`.filter_map(\|a\| a.ok()).find(Requires && Unbound)`): quadratic when Requires is absent or last | NAMES single | no (bool) | none | none | ink placement paragraph scan (`scan_paragraph_candidates`) of story XML |
| `crates/litchi-docx/src/ink/placement.rs:1312` | fallback_image_attribute | checked (explicit) | FAIL-FAST | NAMES r:id (guarded) | yes | O(1) | none | ink `retarget_fallbacks` |
| `crates/litchi-docx/src/ink/placement.rs:1450` | selected_relationship_attribute | checked (explicit) | FAIL-FAST | NAMES r:id (guarded) | yes | O(1) | none (preceded by `xml::element`, 256 cap) | ink `retarget_many` |
| `crates/litchi-docx/src/ink/transaction/drawing_ids.rs:46` | observe | checked | FAIL-FAST | OTHER (every id/xml:id into `HashSet<u32>`) | yes | HashSet | none | ink transaction: drawing-id scan of a story |
| `crates/litchi-docx/src/ink/transaction/durable.rs:2827` | parse_relationships | checked (explicit) | FAIL-FAST | NAMES Id/Type/Target/TargetMode (unknown names refused) | yes | none | none | durable ink patch relationship metadata |
| `crates/litchi-docx/src/ink/xml.rs:118` | element | checked (explicit) | FAIL-FAST | COUNT cap MAX_ATTRIBUTES=256 (Ok only) + expanded names | yes | `names.iter().any` (Vec; n <= 256) | none | shared ink element validator (8 callers) |
| `crates/litchi-docx/src/mail_merge/codec.rs:789` | make_node | checked | FAIL-FAST | ALL; cap MAX_ATTRIBUTES_PER_NODE=256 | yes | none | none | settings mailMerge DOM |
| `crates/litchi-docx/src/math.rs:434` | validate_element_attributes | checked | FAIL-FAST | OTHER (decodes every value) | yes | none | none | OMML extraction |
| `crates/litchi-docx/src/math.rs:510` | root_declares_namespace_binding | checked | FAIL-FAST (returns at first match) | OTHER | yes | none | none | OMML root |
| `crates/litchi-docx/src/modern_comments/codec.rs:1034` | push_node | checked | FAIL-FAST | ALL | yes | `HashSet<(String,String)>` (std), second pass | per-element clone of the parent namespace map | modern-comments parts DOM |
| `crates/litchi-docx/src/namespace.rs:347` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | shared helper (18 callers: settings, content controls, textbox) |
| `crates/litchi-docx/src/namespace.rs:809` | set_direct_property_value | checked | FAIL-FAST | NAMES w:val (guarded) | yes | O(1) | none | Word direct-property reads |
| `crates/litchi-docx/src/numbering/codec.rs:751` | relationship_attribute | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | numbering picture bullets |
| `crates/litchi-docx/src/numbering/codec.rs:836` | effective_ignorable | checked | FAIL-FAST | NAMES mc:Ignorable (guarded) | yes | O(1) | `validation::parse_ignorable` checks each token with `prefixes.iter().any` (Vec): **QUADRATIC** in distinct Ignorable tokens, no token cap. Also `inherited.cloned()` per element | numbering.xml parse |
| `crates/litchi-docx/src/numbering/codec.rs:863` | restart_numbering_after_break | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | `has_ignorable_prefix` linear, called at most once | numbering.xml |
| `crates/litchi-docx/src/numbering/codec.rs:1461` | decoded_attribute_value | checked | FAIL-FAST (returns at first match) | NAMES exact raw key | yes | none | NESTED: `scope_for` (numbering/codec.rs:1365, 1380) calls it once per xmlns declaration from its own raw-attribute loop, so O(decls*n); decls <= 256 (NsReader cap) | numbering `locate_definitions` |
| `crates/litchi-docx/src/numbering/codec.rs:1603` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | numbering.xml (4 direct calls + `required_*` wrappers) |
| `crates/litchi-docx/src/package/package/transfer.rs:589` | relationship_prefixes | checked | FAIL-FAST | OTHER (xmlns into BTreeMap) | yes | none | none | cross-package paragraph transfer (donor XML) |
| `crates/litchi-docx/src/package/package/transfer.rs:640` | stable_namespace_declarations | checked | FAIL-FAST | OTHER (xmlns into BTreeMap) | yes | none | none | same |
| `crates/litchi-docx/src/package/package/transfer.rs:704` | relationship_references | checked | FAIL-FAST | NAMES relationship refs into BTreeSet | yes | none | `reader.resolver().clone()` per event | same |
| `crates/litchi-docx/src/package/package/transfer.rs:824` | rewrite_element | checked | FAIL-FAST | ALL (re-serialized) | yes | BTreeSet `present`, O(log n) | caller clones the resolver per event | same |
| `crates/litchi-docx/src/paragraph/codec/editing.rs:334` | spacing_has_unsupported_attributes | checked | FAIL-FAST (returns true at first unsupported attr) | OTHER | yes | none | none | paragraph spacing edit |
| `crates/litchi-docx/src/paragraph/codec/run.rs:58` | breaks | checked | FAIL-FAST | NAMES type/clear by local name; plain assignment | yes | none | none | public Run API |
| `crates/litchi-docx/src/paragraph/codec/run.rs:462` | vertical_position | checked | **SHORT-CIRCUIT** (`.flatten()`; returns only when val is superscript/subscript): full scan when val is absent, last, or has another value | NAMES val (not stored) | yes | none | none | public Run API over run XML |
| `crates/litchi-docx/src/paragraph/codec/run.rs:509` | font_name | checked | **SHORT-CIRCUIT** (returns at first local `ascii`): full scan when absent or last | NAMES | yes | none | none | public Run API |
| `crates/litchi-docx/src/paragraph/codec/run.rs:556` | font_size | checked | **SHORT-CIRCUIT** (returns at first `val` that parses as u32): full scan when absent, last, or unparsable | NAMES | yes | none | none | public Run API |
| `crates/litchi-docx/src/paragraph/codec/run.rs:611` | get_bool_property | checked | **SHORT-CIRCUIT** (returns at first `val`; if absent, full scan then `Some(true)`) | NAMES | yes | none | none | public Run API (bold/italic/strike) |
| `crates/litchi-docx/src/paragraph/codec/run_properties.rs:40` | update_run_properties | checked | FAIL-FAST (`break` at first val) | NAMES | yes | none | none | run property reads (4 callers) |
| `crates/litchi-docx/src/paragraph/codec/run_properties.rs:276` | run_underline_attribute | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | run underline reads |
| `crates/litchi-docx/src/paragraph/codec/text.rs:565` | observe_namespaces | checked (explicit) | FAIL-FAST | OTHER (byte and binding caps) | yes | none | none | semantic paragraph text |
| `crates/litchi-docx/src/paragraph/codec/text.rs:736` | validate_semantic_attributes | checked (explicit) | FAIL-FAST | OTHER | yes | none | none | same |
| `crates/litchi-docx/src/paragraph/codec/text.rs:761` | validate_semantic_attribute_names | checked (explicit) | FAIL-FAST | OTHER | yes | none | none | same |
| `crates/litchi-docx/src/paragraph/codec/xml.rs:96` | paragraph_attribute | **unchecked** | FAIL-FAST (syntax Err only; returns first local-name match) | NAMES, first wins | yes | none (duplicates silently ignored) | none | paragraph property helpers (6 callers) |
| `crates/litchi-docx/src/paragraph/codec/xml.rs:119` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | paragraph properties (9 callers) |
| `crates/litchi-docx/src/paragraph/collapsed/codec.rs:353` | parse_collapsed | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | w15 collapsed |
| `crates/litchi-docx/src/paragraph/extensions/codec.rs:166` | parse_attributes | checked | FAIL-FAST | NAMES paraId/textId/noSpellErr (guarded) | yes | O(1) | none | w14 paragraph ids |
| `crates/litchi-docx/src/paragraph/package.rs:114` | parse_hyperlink | checked | FAIL-FAST | NAMES (`set_once`) | yes | O(1) | none | direct hyperlink fragment |
| `crates/litchi-docx/src/redact.rs:552` | inspect_owner_element | checked | FAIL-FAST | OTHER | yes | none | `closure.relationships().iter().any` linear over all package relationships per r:* attr: O(attrs x rels) | redaction closure over stories |
| `crates/litchi-docx/src/revision.rs:700` | revision_from_element | checked | FAIL-FAST | COUNT cap `limits.max_attributes` (256) via `enumerate`, checked before `?` (so it counts the one Err item) + NAMES (guarded) | yes | O(1) guards | none | tracked-revision reads |
| `crates/litchi-docx/src/revision/authoring.rs:2311` | parse_metadata | checked | FAIL-FAST | NAMES (guarded; unknown w:* refused) | yes | O(1) | none | tracked-change authoring over source XML |
| `crates/litchi-docx/src/revision/authoring.rs:2409` | parse_range_end_metadata | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | same |
| `crates/litchi-docx/src/revision/authoring.rs:2459` | validate_attributes | checked | FAIL-FAST | COUNT cap `limits.max_attributes` (256 default, <=4,096) (Ok only) | yes | none | none | same; every start element |
| `crates/litchi-docx/src/revision/authoring.rs:2503` | collect_hoisted_attributes | checked (manual `next()`) | FAIL-FAST | ALL (xmlns + MC/xml scope attrs) | yes | `namespace_declarations.iter().any` (Vec; quadratic, bounded by the 256/4,096 cap that `validate_attributes` enforces) + <=5 scope attrs | `raw_tag_attributes` lexical scan once | same |
| `crates/litchi-docx/src/revision/conflict/codec.rs:441` | process_content_directives | checked | FAIL-FAST | OTHER (tokens <= min(max_attributes, 4,096) per directive) | yes | none | none | conflict-markup MCE |
| `crates/litchi-docx/src/revision/conflict/codec.rs:802` | validate_lexical_attributes | checked | FAIL-FAST | COUNT cap max_attributes (64 default) (Ok only) | yes | none | none | conflict markup |
| `crates/litchi-docx/src/revision/conflict/codec.rs:916` | range_end_id | checked | FAIL-FAST | COUNT cap + NAMES id (guarded) | yes | O(1) | `find_attr` once | same |
| `crates/litchi-docx/src/revision/conflict/codec.rs:988` | metadata | checked | FAIL-FAST | COUNT cap + NAMES id/author/date (guarded) | yes | O(1) | `find_attr` rescans the tag lexically for every w:* attr (<= max_attributes=64) | same |
| `crates/litchi-docx/src/revision/conflict/codec.rs:1226` | declares_w14_ignorable | checked | FAIL-FAST (returns at match) | OTHER | yes | none | none | same |
| `crates/litchi-docx/src/revision/conflict/codec.rs:1262` | namespace_declarations | checked | FAIL-FAST | ALL xmlns (cap `max`) | yes | none | none | same |
| `crates/litchi-docx/src/run_effects/codec/mod.rs:331` | node_from_start | checked | FAIL-FAST | ALL | yes | none | none | w14 run effects |
| `crates/litchi-docx/src/run_symbols/codec.rs:343` | parse_symbol | checked | FAIL-FAST | NAMES font/char (guarded) | yes | O(1) | none | w14 symEx |
| `crates/litchi-docx/src/sanitize.rs:717` | external_relationship_id | checked | FAIL-FAST | NAMES r:id (guarded) | yes | O(1) | none (binary search) | external-hyperlink sanitize |
| `crates/litchi-docx/src/sanitize.rs:772` | validate_removable_wrapper_attributes | checked | FAIL-FAST | NAMES (bitset guard; unknown names refused) | yes | bitset | none | same |
| `crates/litchi-docx/src/section/codec.rs:673` | validate_attribute_qnames | checked | FAIL-FAST | OTHER | yes | none | none | section property decode |
| `crates/litchi-docx/src/section/codec.rs:1405` | attributes | checked | FAIL-FAST | ALL relevant attrs (Word/unbound/r:id) into Vec | yes | `result.iter().any(candidate == name)`: **QUADRATIC** on n distinct relevant attrs (for example unprefixed ones); no attribute-count cap | `raw.has_namespace_binding` linear (unknown-prefix case only) | section property decode (`decode_state`: pgSz/pgMar/cols/col/header and footer references; 5 callers) |
| `crates/litchi-docx/src/section/codec.rs:1518` | required_attribute | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | same |
| `crates/litchi-docx/src/section/codec.rs:1574` | root_namespace_bindings | checked | FAIL-FAST | ALL xmlns | yes | none | none | same |
| `crates/litchi-docx/src/section/footnote_columns/codec.rs:526` | parse_extension | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | w15 footnoteColumns |
| `crates/litchi-docx/src/section/footnote_columns/codec.rs:650` | root_ignorable_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | same |
| `crates/litchi-docx/src/section/footnote_columns/package.rs:173` | direct_ignorable | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | same |
| `crates/litchi-docx/src/section/inventory.rs:1216` | validate_namespace_bindings | checked | FAIL-FAST | OTHER | yes | none | none | main-document section inventory |
| `crates/litchi-docx/src/section/inventory.rs:1279` | inspect_reference | checked | FAIL-FAST | NAMES r:id; bindings into Vec | yes | none | `relationship_bindings.iter().any` per r:id attr: quadratic in distinct relationship prefixes (bounded by in-scope bindings) | same |
| `crates/litchi-docx/src/section/inventory.rs:1365` | namespace_declarations | checked | FAIL-FAST | ALL xmlns | yes | none | none | same |
| `crates/litchi-docx/src/section/inventory.rs:1487` | root_declares_prefix | checked | FAIL-FAST (returns at match) | OTHER | yes | none | none | same |
| `crates/litchi-docx/src/section/layout.rs:1044` | word_attribute_info_from_element | checked | FAIL-FAST | ALL Word attrs into `info.prefixes` | yes | none | its consumers `ensure_clear_is_lossless` (layout.rs:1206-1215) and `patch_known_attributes_at` (1436) run a nested `info.prefixes.iter().any` per lexical attr: quadratic, but only attrs with known local names reach it | section layout edit |
| `crates/litchi-docx/src/section/layout.rs:1146` | rewrite_known_children | NOT-QX (`KnownChild::attributes()` returns `&'static [&[u8]]`) | n/a | n/a | n/a | n/a | n/a | n/a |
| `crates/litchi-docx/src/section/layout.rs:1196` | ensure_clear_is_lossless | NOT-QX (`KnownChild::attributes()`) | n/a | n/a | n/a | n/a | n/a | n/a |
| `crates/litchi-docx/src/section/layout.rs:1212` | ensure_clear_is_lossless | NOT-QX (`KnownChild::attributes()`) | n/a | n/a | n/a | n/a | n/a | n/a |
| `crates/litchi-docx/src/section/layout.rs:1397` | patch_columns | NOT-QX (`KnownChild::Columns.attributes()`) | n/a | n/a | n/a | n/a | n/a | n/a |
| `crates/litchi-docx/src/settings/codec.rs:583` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | settings.xml (11 callers) |
| `crates/litchi-docx/src/settings/document/codec.rs:654` | relationship_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | settings attachedTemplate |
| `crates/litchi-docx/src/settings/document/codec.rs:1426` | capture_settings_root | checked | FAIL-FAST (`break` at first relationship binding) | OTHER | yes | none | none | settings root |
| `crates/litchi-docx/src/settings/extensions/codec.rs:868` | optional_attribute | checked | FAIL-FAST | NAMES val (guarded; unknown names refused) | yes | O(1) | none | settings extensions |
| `crates/litchi-docx/src/settings/extensions/codec.rs:976` | make_self_contained | checked | FAIL-FAST | OTHER (xmlns into HashSet) | yes | HashSet (std) | none | same |
| `crates/litchi-docx/src/smart_tag.rs:177` | optional_attribute | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | smart tags (7 callers) |
| `crates/litchi-docx/src/smartart.rs:518` | rel_ids | checked (explicit) | FAIL-FAST | NAMES r:dm/lo/qs/cs; plain assignment (overwrites across prefixes) | yes | none | none | SmartArt graph scan |
| `crates/litchi-docx/src/smartart.rs:547` | attribute | checked (explicit) | FAIL-FAST (returns at first match) | NAMES | yes | none | none | same |
| `crates/litchi-docx/src/source_backed.rs:2759` | validate_source_section_element | checked | FAIL-FAST | OTHER | yes | none | none | source-backed section inventory |
| `crates/litchi-docx/src/source_backed.rs:2790` | ensure_source_mce_element | checked | FAIL-FAST | OTHER | yes | none | none | source-backed document read and section inventory (every element) |
| `crates/litchi-docx/src/source_backed/document_policy.rs:508` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | changed-document publication (settings policy) |
| `crates/litchi-docx/src/source_backed/document_policy.rs:552` | reject_settings_mce | checked | FAIL-FAST | OTHER | yes | none | none (prefix lookups are charged to the execution budget) | same |
| `crates/litchi-docx/src/source_backed/paragraph_copy.rs:1808` | validate_attributes | checked | FAIL-FAST | OTHER (closed allow-list) | yes | none | none | source-backed paragraph copy |
| `crates/litchi-docx/src/source_backed/story_text.rs:2384` | for_element | checked (explicit) | FAIL-FAST | ALL xmlns | yes | none | per-element clone of the parent bindings Vec (<= max_bindings) | source-backed story text |
| `crates/litchi-docx/src/source_backed/story_text.rs:2464` | for_element | checked (explicit) | FAIL-FAST | OTHER | yes | none | `resolve_prefix`/`resolve_attribute` linear over the bindings Vec per attr (<= max_bindings) | same |
| `crates/litchi-docx/src/source_backed/story_text.rs:2711` | word_attribute | checked (explicit) | FAIL-FAST (returns `Some(Err)`) | NAMES (guarded) | yes (`Option<Result>`) | O(1) | linear resolve per matching attr | same |
| `crates/litchi-docx/src/source_backed/story_text.rs:4372` | reject_word_dialect_attributes | checked (explicit) | FAIL-FAST | OTHER | yes | none | linear `resolve_prefix` per prefixed attr | same |
| `crates/litchi-docx/src/source_backed/tail_append.rs:1711` | settings_namespace_declarations | checked | FAIL-FAST | COUNT xmlns (Ok only; returned, not capped here) | yes | none | none | tail-append scan of settings.xml |
| `crates/litchi-docx/src/source_backed/tail_append.rs:1767` | settings_event_owned_bytes | checked | FAIL-FAST | COUNT attrs (returned) | yes | none | none | same |
| `crates/litchi-docx/src/source_backed/tail_append.rs:1802` | settings_mce_directive_facts | checked | FAIL-FAST | OTHER (MCE directive tokens) | yes | none | `resolver.bindings().find_map(..)` per token. quick-xml's `NamespaceBindingsIter` rescans later bindings, so a miss costs O(B^2) and the total is O(tokens x B^2); tokens uncapped here | same |
| `crates/litchi-docx/src/source_backed/tail_append.rs:3359` | validate_xml_declaration | checked (explicit) | FAIL-FAST | NAMES (closed state machine) | yes | state machine | none | XML declaration of scanned parts |
| `crates/litchi-docx/src/source_backed/tail_append.rs:3490` | validate_opaque_attributes | checked | **COUNT**: `element.attributes().count()` runs first, walks every item and counts Err items too, so it is a tolerant full traversal | COUNT (only feeds `try_reserve`) | yes | n/a | none beyond this pre-count, which is itself the Theta(D*R) pass | tail-append source document.xml: every element of the opaque final `w:sectPr` (4 callers) |
| `crates/litchi-docx/src/source_backed/tail_append.rs:3495` | validate_opaque_attributes | checked | FAIL-FAST | ALL expanded names into Vec | yes | `sort_unstable` + `windows(2)`, O(n log n) | none | same |
| `crates/litchi-docx/src/source_backed/tail_append.rs:3540` | validate_attributes | checked | FAIL-FAST | OTHER (closed) | yes | none | none | tail-append plain story |
| `crates/litchi-docx/src/story_hyperlinks.rs:1647` | inspect_story_element | checked | FAIL-FAST | OTHER | yes | none | none (binary search) | story hyperlink inventory |
| `crates/litchi-docx/src/story_hyperlinks.rs:1679` | has_external_hyperlink_id | checked | FAIL-FAST (returns at first r:id) | NAMES | yes | none | none | same |
| `crates/litchi-docx/src/story_hyperlinks.rs:1861` | validate_dialect_attributes | checked | FAIL-FAST | OTHER | yes | none | none | same |
| `crates/litchi-docx/src/styles.rs:329` | ensure_styles_loaded | checked | **TOLERANT** (`.flatten()`; decode errors are also ignored) | NAMES type/styleId/default/customStyle by local name; OVERWRITE | yes | none | none | styles.xml lazy parse on the first style query. `mce::process_part` returns the raw bytes when there is no MCE namespace, so no earlier duplicate check runs |
| `crates/litchi-docx/src/styles.rs:430` | ensure_styles_loaded | checked | **TOLERANT** (`.flatten()`) | NAMES name@val; OVERWRITE | yes | none | none | same |
| `crates/litchi-docx/src/styles.rs:443` | ensure_styles_loaded | checked | **TOLERANT** (`.flatten()`) | NAMES basedOn@val; OVERWRITE | yes | none | none | same |
| `crates/litchi-docx/src/styles.rs:456` | ensure_styles_loaded | checked | **TOLERANT** (`.flatten()`) | NAMES uiPriority@val; OVERWRITE | yes | none | none | same |
| `crates/litchi-docx/src/styles.rs:680` | required_style_value | **unchecked** | FAIL-FAST (syntax Err only; returns first local `val`) | NAMES, first wins | yes | none | none | styles numId/ilvl/outlineLvl (3 callers) |
| `crates/litchi-docx/src/styles/effects.rs:2498` | relationship_element_id | checked (explicit) | FAIL-FAST (returns at first `Id`) | NAMES | yes | none | none | stylesWithEffects relationships |
| `crates/litchi-docx/src/styles/effects.rs:3047` | root_namespace | **unchecked** | FAIL-FAST (syntax Err only; first match) | NAMES, first wins | yes | none | per attr, a `concat` allocation and `normalized_value` run before the match test (linear) | stylesWithEffects root |
| `crates/litchi-docx/src/table.rs:130` | word_cell_property_value | checked | FAIL-FAST | NAMES val by local name; plain assignment | yes | none | none | table cell property reads |
| `crates/litchi-docx/src/validation.rs:1097` | is_valid_element_attributes | checked (explicit) | FAIL-FAST (`.all` returns false at first Err) | OTHER | no (bool; caller refuses) | none | none | package validation of the visible main document |
| `crates/litchi-docx/src/validation.rs:1109` | has_unknown_attribute_namespace | checked (explicit) | **SHORT-CIRCUIT** (`.any(is_ok_and ..)`; an Err gives false and iteration continues): full scan when nothing matches | OTHER | no (bool; caller refuses) | none | none | same. Both callers evaluate it only after 1097 returned true (`\|\|`), so the tag has no Err items and the cost is linear in practice |
| `crates/litchi-docx/src/variables/codec.rs:922` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | settings docVars (3 callers) |
| `crates/litchi-docx/src/web/codec/relationship.rs:69` | required_relationship_id | checked | FAIL-FAST | NAMES r:id (guarded) | yes | O(1) | none | webSettings frame source |
| `crates/litchi-docx/src/web/codec/rewrite.rs:581` | attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | webSettings rewrite |
| `crates/litchi-docx/src/web/mod.rs:95` | word_attribute_value | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | webSettings (9 callers) |
| `crates/litchi-docx/src/writer/doc/codec.rs:190` | element_preserves_space | checked | FAIL-FAST | NAMES xml:space; plain assignment | yes | none | none | compaction of changed document.xml at save |
| `crates/litchi-docx/src/writer/doc/codec.rs:219` | write_compact_start | checked | FAIL-FAST | ALL (re-serialized) | yes | none | none | same |
| `crates/litchi-docx/src/writer/doc/fusion_tests_0722.rs:461` | frozen_alt_relationship | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | **TEST** (`#[cfg(test)] mod fusion_tests_0722`) |
| `crates/litchi-docx/src/writer/doc/fusion_tests_0722.rs:500` | frozen_alt_parse_on_off | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | **TEST** |
| `crates/litchi-docx/src/writer/section/codec/xml.rs:1380` | root_metadata | checked | FAIL-FAST | ALL | yes | 3 HashSets (std) | none | writer `SectionProperties::from_xml` (sectPr) |
| `crates/litchi-docx/src/writer/section/codec/xml.rs:1484` | is_word_element_at | checked | **SHORT-CIRCUIT** (`.any(\|a\| a.ok().is_some_and(xmlns==""))`): the target `xmlns=""` is normally absent, so it scans everything | OTHER | no (bool; caller returns Result) | none | none | writer `direct_children`, for each unprefixed descendant when the root has a default Word namespace. It runs **before** the FAIL-FAST `validate_direct_child_attributes` on the same tag |
| `crates/litchi-docx/src/writer/section/codec/xml.rs:1986` | decode_attributes | checked | FAIL-FAST | ALL | yes | 2 HashSets (std) | `metadata.has_binding` linear (unknown-prefix case) | same |
| `crates/litchi-docx/src/writer/section/validation.rs:162` | validate_attributes | checked | FAIL-FAST | OTHER | yes | 2 HashSets (std) | `reader.resolver().clone()` per element | writer header/footer XML validation |
| `crates/litchi-docx/src/writer/watermark.rs:511` | unqualified_attribute | checked | FAIL-FAST | NAMES (guarded) | yes | O(1) | none | watermark header VML scan (9 callers) |
| `crates/litchi-docx/src/writer/watermark.rs:931` | apply_image_data | checked | FAIL-FAST | NAMES r:id; plain assignment (overwrites across prefixes) | yes | none | none | same |

## Helper functions that wrap attribute iteration

Every caller inherits the helper's mode. Counts are call sites in the stated scope (grep, definition line excluded, method calls
like `.attributes(` excluded). All paths are under `crates/litchi-docx/src/`.

### Non-fail-fast helpers (the ones that matter for hardening)

| helper | mode | call sites |
|---|---|---|
| `drawing/codec.rs:371 inert_attribute` | TOLERANT, full scan, last wins | 3 in drawing/codec.rs: name, descr, and `number_attribute` |
| `drawing/codec.rs:359 number_attribute` (wraps `inert_attribute`) | TOLERANT | 6 in drawing/codec.rs (cx, cy, x, y, cx, cy), so 8 effective call sites in `drawing::parse` |
| `paragraph/codec/run.rs:596 get_bool_property` | SHORT-CIRCUIT | 3 (run.rs:144 `b`, 157 `i`, 265 `strike`) |
| `ink/placement.rs:1213 choice_requires_supported` | SHORT-CIRCUIT | 1 (ink/placement.rs:1070) |
| `validation.rs:1105 has_unknown_attribute_namespace` | SHORT-CIRCUIT (guarded by caller) | 2 (validation.rs:1197, 1246; each only after `is_valid_element_attributes` returns true) |
| `writer/section/codec/xml.rs:1474 is_word_element_at` | SHORT-CIRCUIT | 1 (xml.rs:1159) |
| `source_backed/tail_append.rs:3485 validate_opaque_attributes` | tolerant COUNT, then FAIL-FAST | 4 (tail_append.rs:2950, 2973, 3043, 3120) |
| `document/transaction.rs:6149 event_namespace_binding_count` | COUNT, unchecked (linear) | 1 (transaction.rs:6273) |

### Fail-fast single-attribute lookups and validators

The `word_attribute_value` family has 7 separate definitions, all checked, fail-fast and guarded:

| definition | call sites |
|---|---|
| `namespace.rs:340` (pub(crate)) | 18: settings/document/codec.rs 11, content_control.rs 5, textbox.rs 2 |
| `paragraph/codec/xml.rs:111` (pub(super)) | 9: paragraph_properties.rs 8, run_contents.rs 1 |
| `web/mod.rs:88` (pub(super)) | 9, all in web/codec/xml.rs |
| `settings/codec.rs:576` | 11 |
| `numbering/codec.rs:1596` | 4 direct. Also wrapped by `required_string` (8 calls), which is wrapped by `required_u32` (10) and `required_i64` (2); `required_u32` is wrapped by `required_level` (6) |
| `variables/codec.rs:915` | 3 |
| `source_backed/document_policy.rs:501` | 1 |

Other fail-fast helpers:
- Unchecked, return the first match:
  - `paragraph/codec/xml.rs:91 paragraph_attribute`: 6 calls in paragraph/codec/.
  - `styles.rs:676 required_style_value`: 3.
  - `styles/effects.rs:3040 root_namespace`: 2.
- Numbering:
  - `numbering/codec.rs:1456 decoded_attribute_value`: 2. Both calls sit inside `scope_for` loops over raw attributes, so the pattern is nested.
  - `numbering/codec.rs:744 relationship_attribute`: 1.
- Drawing:
  - `drawing/codec.rs:385 strict_attribute`: 1.
  - `drawing/codec.rs:324 parse_anchor_id`: 2.
  - `drawing/source.rs:2840 relationship_attribute`: 2.
  - `drawing/source.rs:2881 unqualified_attribute`: 3.
  - `drawing/source.rs:2560 relationship_dialect_for_prefix`: 1, but it runs once per attribute inside the 2525 loop.
  - `drawing/validation.rs:38 parse_word2010_anchor_id`: 2 crate-wide.
- Content controls:
  - `content_control.rs:1568 extension_attribute_value`: 2.
  - `content_control.rs:1604 exact_extension_attribute_value`: 2.
  - `content_control/snapshot.rs:901 exact_attribute`: 4, all per element.
- Ink:
  - `ink/xml.rs:107 element` (validator with a 256 cap): 8 production calls (ink/codec.rs 1, ink/host.rs 4, ink/placement.rs 3) plus 6 test calls.
  - `ink/package.rs:1227 profile_attr`: 6.
  - `ink/package.rs:1207 xml_id`: 2.
  - `ink/codec.rs:1067 graphic_data_kind` and `ink/host.rs:979 graphic_data_kind`: 1 each. Both return None at the first Err.
- Settings, smart tags, SmartArt, web, watermark:
  - `settings/document/codec.rs:647 relationship_attribute_value`: 1.
  - `settings/extensions/codec.rs:860 optional_attribute`: 3.
  - `smart_tag.rs:171 optional_attribute`: 7.
  - `smartart.rs:546 attribute`: 2.
  - `web/codec/rewrite.rs:561 attribute_value`: 1.
  - `web/codec/relationship.rs:59 required_relationship_id`: 1.
  - `writer/watermark.rs:504 unqualified_attribute`: 9.
- Sections:
  - `section/codec.rs:1364 attributes`: 5. It is fail-fast, but all 5 callers inherit its uncapped QUADRATIC own-duplicate check.
  - `section/codec.rs:1480 required_attribute`: 1.
- Paragraph and run properties:
  - `paragraph/codec/run_properties.rs:268 run_underline_attribute`: 5.
  - `paragraph/codec/run_properties.rs:27 update_run_properties`: 4.
  - `source_backed/story_text.rs:2704 word_attribute`: 3.
- Validation, math, altChunk:
  - `validation.rs:1096 is_valid_element_attributes`: 2.
  - `math.rs:430 validate_element_attributes`: 2.
  - `math.rs:506 root_declares_namespace_binding`: 2.
  - `alt/codec.rs:312 relationship`: 2.
  - `alt/codec.rs:351 parse_on_off`: 1.

## Summary

### Counts

175 grep hits: 168 production quick-xml sites, 3 TEST sites and 4 NOT-QX sites.

| class | production sites |
|---|---|
| FAIL-FAST | 148 (5 of them unchecked) |
| TOLERANT | 9 (all checked) |
| SHORT-CIRCUIT | 8 (all checked; 1 guarded by its caller) |
| COUNT | 2 (1 checked, 1 unchecked) |
| n/a (`size_hint` only) | 1 |

- Checks: 162 production sites are checked and 6 are unchecked. TEST adds 1 unchecked and 2 checked sites.
- TEST sites: `document/transaction.rs:6525` (COUNT, unchecked), `writer/doc/fusion_tests_0722.rs:461` and `:500` (FAIL-FAST).
- NOT-QX sites: `section/layout.rs:1146`, `1196`, `1212`, `1397` (`KnownChild::attributes()`).
- 17 production sites are exposed to Theta(D*R) on plain duplicates: 9 TOLERANT, 7 SHORT-CIRCUIT and 1 COUNT.
  - `validation.rs:1109` is not exposed because its callers guard it.
  - `document/transaction.rs:6156` is not exposed because it is unchecked and therefore linear.

### TOLERANT production sites

All paths below are under `crates/litchi-docx/src/`.
- `bookmark.rs:85`: `Package::bookmarks()`; id/name overwritten.
- `drawing/codec.rs:373`: `inert_attribute`, 8 calls per drawing.
- `image.rs:308`: inline image cx/cy; overwritten.
- `image.rs:327`: inline image name/descr; overwritten.
- `image.rs:349`: blip `*:embed`; last one wins.
- `styles.rs:329`: style type/styleId/default/customStyle attributes.
- `styles.rs:430`: style name@val; overwritten.
- `styles.rs:443`: style basedOn@val; overwritten.
- `styles.rs:456`: style uiPriority@val; overwritten.

### SHORT-CIRCUIT production sites

Each is quadratic when the searched name is absent or last.
- `document/transaction.rs:8771`: finds fldCharType; absent means refused.
- `ink/placement.rs:1215`: finds an unbound `Requires` on mc:Choice.
- `paragraph/codec/run.rs:462`: `vertical_position`; returns only on superscript/subscript.
- `paragraph/codec/run.rs:509`: `font_name`; finds `ascii`.
- `paragraph/codec/run.rs:556`: `font_size`; finds a numeric `val`.
- `paragraph/codec/run.rs:611`: `get_bool_property`; `val` is usually absent.
- `validation.rs:1109`: unknown namespace. Its caller guards it, so it is linear in practice.
- `writer/section/codec/xml.rs:1484`: looks for `xmlns=""`, which is usually absent. It runs before the fail-fast check.

### COUNT production sites

- `source_backed/tail_append.rs:3490`: `.count()` runs first and counts Err items, so it is quadratic.
- `document/transaction.rs:6156`: unchecked `filter_map` count, so it is linear.

### QUADRATIC own-duplicate checks

Uncapped, or effectively uncapped:
- `chart/codec.rs:380`: `values.iter().any` over expanded names. The only caps are 750,000 attributes per part and 16 MiB of attribute bytes.
- `section/codec.rs:1405`: `result.iter().any` over local names, with no count cap. 5 callers inherit it.
- Related: `numbering/validation.rs:50 parse_ignorable` checks each mc:Ignorable token with `prefixes.iter().any` and has no token cap. It is reached from `numbering/codec.rs:836` (and from 1381 outside the site list).

Bounded by small caps:
- `drawing/source.rs:2662` (n <= 256).
- `ink/xml.rs:118` (n <= 256).
- `revision/authoring.rs:2503`: `namespace_declarations`, n <= 256 by default, 4,096 at most.
- `document/transaction.rs:7678` (at most 5 entries, O(1)).
- `content_control/snapshot.rs:731`: `local_ignorable` dedupe, up to 65,536 tokens, with no attribute-level duplicate detection because the site is unchecked.

### Other super-linear patterns

Unbounded or large:
- `source_backed/tail_append.rs:1802` calls `resolver.bindings().find_map` per MCE token. quick-xml's bindings iterator is O(B^2) per miss, so the total is O(tokens x B^2) with tokens uncapped.
- `redact.rs:552` scans every package relationship for each r:* attribute, so O(attrs x rels).
- `document/transaction.rs:8771`: its caller reparses the paragraph from the start for each run, so O(runs x bytes) with runs <= 4,096.

Nested re-iteration, bounded:
- `drawing/source.rs:2525` calls `relationship_dialect_for_prefix` for each *:embed/*:link attribute (n <= 256).
- `numbering/codec.rs:1461` is called once per xmlns declaration from `scope_for` (declarations <= 256 x n).

Lexical `find_attr` rescan of the tag per matching attribute:
- `content_control/snapshot.rs:914` (at most 1,024 matches).
- `revision/conflict/codec.rs:988` (at most 64 matches).

Linear Vec scans per attribute:
- `section/inventory.rs:1279` over `relationship_bindings`.
- `section/layout.rs:1044`: its consumers run nested `info.prefixes` scans.
- `source_backed/story_text.rs:2464`, `2711` and `4372` resolve against a bindings Vec (<= max_bindings).

Per-element or per-event scope clones:
- Per element: `font/codec.rs:148`, `modern_comments/codec.rs:1034`, `source_backed/story_text.rs:2384`, `numbering/codec.rs:836`.
- `resolver().clone()` per event or element: `package/package/transfer.rs:704` and `:824` (in the caller), `complex_field_marker`, numbering `locate_definitions`, `writer/section/validation.rs:162`.

Crate-wide, every resolving site pays a linear `resolve_attribute` over the in-scope bindings. NsReader caps declarations at 256 per element by default.

### Semantic notes for hardening

- The TOLERANT `.flatten()` sites also silently skip syntax errors such as ExpectedEq or unquoted values, not only duplicates.
- These sites store values by plain assignment:
  - `bookmark.rs:85`, `image.rs:308`, `327` and `349`, `styles.rs:329`, `430`, `443` and `456`, `drawing/codec.rs:373`.
  - Switching them to unchecked would change which value wins from first to last. Refusing the whole call would change it from silent to error. `drawing/codec.rs:373`'s helper returns Option, so its refusal would have to go through `number_attribute`, which returns Result.
- Local-name matching means distinct prefixes such as `a:val` and `b:val` already overwrite each other at several FAIL-FAST sites. This is not a quick-xml duplicate, so the checked iterator does not catch it. The sites are:
  - `comment.rs:273`, `field/codec.rs:86`, `hyperlink.rs:190`, `table.rs:130`, `smartart.rs:518`, `writer/watermark.rs:931`, `image.rs:349`.
