# 0764 survey: quick-xml attribute-iterator call sites in the "misc" crates

Worktree `/home/zhuhe/code/litchi-worktrees/0764-xml-attribute-dos-hardening` at `1d1044e3ac`.
Crates: litchi-drawingml, litchi-formula, litchi-xlsb, litchi-spreadsheet-drawing, litchi-xldm,
litchi-ppt, litchi-ole-common, litchi-imgconv, litchi-crypto, litchi-sign (all `src/**/*.rs`).
Search: `\.(html_)?attributes\(\)` (catches split method chains), plus `Attributes::new|html`,
`html_attributes`, `attributes_raw`, `try_get_attribute`. 90 hits: 88 quick-xml `BytesStart`
iterators, 2 NOT-QX. No `html_attributes`/`Attributes::new` anywhere in scope. litchi-imgconv has
0 sites (its only quick-xml use is a test-only `Reader` in `svg.rs`). No quick-xml site is inside
a `#[cfg(test)]` module (the `#[cfg(test)]` at `litchi-xldm/src/codec.rs:242` covers one fn only).

quick-xml 0.41.0 semantics confirmed in the registry source: `check_for_duplicates`
(`events/attributes.rs:1135`) scans linearly below 32 keys and uses the hash prefilter plus a
linear `keys.iter().find` on every hit at or above 32; duplicates are never pushed to `keys`; on
`Err(Duplicated)` the state becomes `SkipEqValue` and iteration continues. `NamespaceResolver::push`
(`name.rs:708`, run by every NsReader event) is `with_checks(false)` and `break`s at the first Err,
with a default cap of 256 declarations per element (`DEFAULT_MAX_DECLARATIONS_PER_ELEMENT`).
`BytesStart::try_get_attribute` is `with_checks(false)` + `a?` (linear, fail-fast).

Legend: FF = FAIL-FAST, SC = SHORT-CIRCUIT. "checked (default)" = plain `.attributes()`;
"checked (explicit)" = `.with_checks(true)`. "GUARDED" = the tolerant site only ever sees an
element that already passed a fail-fast checked pass (with a count cap) in every caller path.

## Site table

| # | file:line | fn | checks | error mode | consumption | can-refuse | own-duplicate-check | other super-linear | reachability |
|---|---|---|---|---|---|---|---|---|---|
| 1 | crates/litchi-crypto/src/labels.rs:444 | parse_label | checked (explicit) | FF (`map_err(..)?`) | n/a | yes | per-name `Option::replace` slots, O(1) | none | sensitivity-label XML (LabelInfo) parse at open |
| 2 | crates/litchi-crypto/src/labels.rs:620 | validate_root | checked (explicit) | FF | n/a | yes | none (HashSet of prefixes, not a dup check) | none (1 extra fail-fast `optional_attribute` pass) | label XML root |
| 3 | crates/litchi-crypto/src/labels.rs:750 | optional_attribute (helper) | checked (explicit) | FF | n/a | yes | `Option::replace` guard | none | label XML (3 callers + `required_attribute` x3) |
| 4 | crates/litchi-crypto/src/labels.rs:766 | reject_non_namespace_attributes | checked (explicit) | FF (stops at first non-xmlns) | n/a | yes | none | none | label XML container elements |
| 5 | crates/litchi-crypto/src/ooxml/agile.rs:1590 | exact_attributes | checked (explicit) | FF; doc-wide counter `max_xml_attributes` (default 4096) incremented per item BEFORE the `?` (counts the Err item) | n/a | yes | u16 bitset over <=16 allowed names | `allowed.iter().position` (<=16, constant) | Agile EncryptionInfo XML, parsed before decryption at open (untrusted) |
| 6 | crates/litchi-drawingml/src/blip.rs:52 | read_embed | checked (default) | FF; returns at first `embed` | n/a | yes | none | none | public `read_embed`/`find_first_embed`; only test callers in repo |
| 7 | crates/litchi-drawingml/src/chart/extension/formatcode2.rs:417 | preflight_namespace_limits | checked (default) | FF; xmlns-only counter, cap 64 (after `?`) | n/a | yes | none | none | formatcode2 chart-extension XML (<=1 MiB) preflight |
| 8 | crates/litchi-drawingml/src/chart/extension/formatcode2.rs:524 | local_namespace_prefixes | checked (default) | FF; prefix cap 64 | n/a | yes | none | none in loop (caller `checked_inherited_bindings` does `local_prefixes.iter().any` per binding, <=64x64) | formatcode2 attribute-fragment start tag |
| 9 | crates/litchi-drawingml/src/chart/extension/formatcode2.rs:677 | attribute_projection | checked (default) | FF | n/a | yes | `selected.is_some()` guard | per-attr `resolve_attribute` (NsReader decl cap 64) | formatcode2 host start tag |
| 10 | crates/litchi-drawingml/src/chart/extension/formatcode2.rs:1229 | only_namespace_attributes | checked (default) | FF; `Ok(false)` at first non-xmlns | n/a | yes | none | none | formatcode2 element |
| 11 | crates/litchi-drawingml/src/chart/extension/formatcode2.rs:1271 | validate_attributes | checked (default) | FF; count-first cap 64 (counter before `?`, counts Err items) | n/a | yes | `HashSet<&[u8]>` raw names + `HashSet<(ns,local)>` (std RandomState/SipHash) | per-attr `resolve_attribute` | formatcode2 elements |
| 12 | crates/litchi-drawingml/src/chart/extension/formatcode2.rs:1428 | validate_declaration | checked (default) | FF | n/a | yes | `seen_*` bools | none | `<?xml ..?>` pseudo-attributes (BytesStart::from_content) |
| 13 | crates/litchi-drawingml/src/chart/model.rs:37 | validate_chart_xml_fragment | checked (default) | FF | n/a | yes | none | per-attr `resolve_attribute` | `ShapeProperties/TextProperties/ExtensionList::from_xml` (validates every chart-reader capture after #16 ran) |
| 14 | crates/litchi-drawingml/src/chart/reader/codec/validation.rs:422 | get_attr (helper) | checked (default) | SC: `.filter_map(Result::ok).find(key == name)`; quadratic when the name is absent (typical: optional `val`) or sits after the duplicate run | NAMES: first match wins; duplicates are Err and dropped, no overwrite | no (returns Option; all callers return Result) | none | called several times per element (e.g. `parse_number_format` = `try_get_attribute` + `get_attr`) | chart part open (DOCX `litchi-docx/src/chart/codec.rs:241`, XLSX `litchi-spreadsheet-drawing/src/chart/codec.rs:30`); reachable with duplicates (see summary: MCE pre-pass skipped without MCE namespace; chart reader has no checked pass on non-root elements) |
| 15 | crates/litchi-drawingml/src/chart/reader/model.rs:67 | relationship_attribute_value | checked (default) | FF | n/a | yes | `value.is_some()` guard | `resolve_attribute` per local-name match | chart reader r:id lookups |
| 16 | crates/litchi-drawingml/src/chart/reader/model.rs:103 | make_fragment_root_self_contained | checked (default) | TOLERANT: `.filter_map(Result::ok)` over the whole tag | ALL: collects every Ok key into `existing_names` (only used to skip root xmlns already present) | no (returns `BytesStart<'static>`; callers `capture_fragment`/`capture_empty_fragment` return Result) | none | `existing_names.iter().any` per chartSpace-root namespace decl: O(K*n), K <= ~259 | chart part open: root of every captured c:spPr/c:txPr/c:extLst etc. (51 `capture_fragment` + 51 `capture_empty_fragment` sites); the later fail-fast #13 refuses the fragment only after this cost is paid |
| 17 | crates/litchi-drawingml/src/chart/reader/model.rs:206 | read_event_into | checked (default) | FF | n/a | yes | none (Vec push; dups impossible) | `root_namespace_attributes.iter().any` for 3 fixed names | chartSpace root only (depth 0) |
| 18 | crates/litchi-drawingml/src/chart/style/codec.rs:662 | make_node | checked (explicit) | FF; non-xmlns cap 64 (checked before push) | n/a | yes | `attributes.iter().any` on expanded name: QUADRATIC but capped (<=64^2) | per-attr `resolve_attribute`; xmlns decls not capped in loop (NsReader default 256/element) | chart style part (cs:chartStyle) parse |
| 19 | crates/litchi-drawingml/src/color/codec.rs:396 | attributes (helper) | checked (default) | FF; returns `Ok(None)` at first prefixed/unknown attr | n/a | yes | `values.iter().any`: bounded by the tiny `allowed` list | `allowed.contains` (constant) | DrawingML color element parse (2 callers) |
| 20 | crates/litchi-drawingml/src/diagram/data/codec.rs:745 | attributes (helper) | checked (explicit) | FF | n/a | yes | none (callers assign by name; dups impossible) | none | SmartArt data-model parse (3 callers) |
| 21 | crates/litchi-drawingml/src/diagram/definition.rs:193 | attribute (helper) | checked (explicit) | FF; returns at first local-name match | n/a | yes | none | none | SmartArt layout-definition parse (5 callers) |
| 22 | crates/litchi-drawingml/src/ink/actions.rs:1091 | validate_profile_attributes | checked (default) | FF | n/a | yes | none | none | ink actions profile parse (8 callers) |
| 23 | crates/litchi-drawingml/src/ink/actions.rs:1721 | optional_attr | checked (default) | FF | n/a | yes | `result.is_some()` guard | none | ink actions (3 callers) |
| 24 | crates/litchi-drawingml/src/ink/actions.rs:1751 | optional_xml_id | checked (default) | FF | n/a | yes | `result.is_some()` guard | none | ink actions (4 callers) |
| 25 | crates/litchi-drawingml/src/ink/actions.rs:1884 | required_attr | checked (default) | FF | n/a | yes | `result.is_some()` guard | none | ink actions (12 callers) |
| 26 | crates/litchi-drawingml/src/ink/actions.rs:2074 | validate_element | checked (default) | FF; cap 256 (`attribute_keys.len()` after `?`) | n/a | yes | pairwise `expanded_attribute_names_equal` against all earlier keys: QUADRATIC but capped (<=256^2/2 pairs) | each pair calls `resolve_attribute` twice (O(in-scope bindings)) before comparing locals | ink actions part parse; runs first on every element (5 callers) |
| 27 | crates/litchi-drawingml/src/ink/actions.rs:2325 | has_namespace_declaration | checked (default) | SC: `.any(..)` with `let Ok(..) else { return false }`; the searched `xmlns:<p>` is normally absent (only called when the prefix is Unknown), so full scan | OTHER (existence test) | no (bool) | none | 1-2 calls per element | GUARDED: all 3 callers (`require_root`, `is_action_namespace`, `validate_unknown_prefix`) run after #26 (FF, cap 256) on the same element |
| 28 | crates/litchi-drawingml/src/ink/actions_edit.rs:1798 | payload_has_namespace_declaration | checked (default) | SC: same shape as #27; absent name, full scan | OTHER (existence test) | no (bool) | none | called for the payload root (:1767), per element (:2014 via `payload_expanded_element_name`) and per Unknown-prefix attribute inside #29's loop (:2115) | NOT GUARDED: ink-action draft `OpaquePayload` (caller-authored, <=16 MiB `MAX_SOURCE_BYTES`) when `allow_inherited_namespace` (definitions payload: always) |
| 29 | crates/litchi-drawingml/src/ink/actions_edit.rs:2059 | validate_payload_attributes_with_namespaces | checked (default) | FF | n/a | yes | `HashSet<PayloadExpandedName>` (std SipHash) | NESTED RE-SCAN: every `iact:`/`inkml:` attribute whose prefix is Unknown calls #28 (full tolerant scan of the whole tag): Theta(n^2) with NO duplicates, Theta(D^2*R) with duplicates; no attribute cap before this pass (cap 256 only in pass 3, #30) | ink-action draft payload validation pass 2 (`validate_payload_namespace_grammar`) |
| 30 | crates/litchi-drawingml/src/ink/actions_edit.rs:2245 | validate_payload_element | checked (default) | FF; cap 256 (count after `?`) | n/a | yes | none | none | payload pass 3 (runs after #28/#29) |
| 31 | crates/litchi-drawingml/src/ink/actions_edit.rs:4479 | check_profile_attributes | UNCHECKED (`with_checks(false)`) | FF (malformed only) | n/a | yes | none (bounds sizes/characters only; names not stored) | none | `validate_profile_limits` re-scan of an already-parsed profile source |
| 32 | crates/litchi-drawingml/src/ink/actions_edit.rs:5253 | index_opaque_payload | checked (default) | FF | n/a | yes | none | `push_reference`/HashMap, O(1) amortized | xml:id index of retained profile payloads (edit) |
| 33 | crates/litchi-drawingml/src/ink/actions_edit.rs:5432 | index_draft_opaque_references | checked (default) | FF | n/a | yes | none | HashMap entry, O(1) | draft payload references (edit) |
| 34 | crates/litchi-drawingml/src/ink/actions_edit.rs:5560 | collect_opaque_ids | checked (default) | FF | n/a | yes | none (later HashSet) | none | draft payload xml:ids (edit) |
| 35 | crates/litchi-drawingml/src/ink/authoring.rs:860 | preflight_attributes | checked (explicit) | FF; cap `max_attributes_per_element` (default 256, count after `?`) | n/a | yes | none | none | canonical InkML draft import preflight |
| 36 | crates/litchi-drawingml/src/ink/authoring.rs:1295 | required_attribute | checked (explicit) | FF | n/a | yes | `value.is_some()` guard | none | canonical InkML import (7 callers) |
| 37 | crates/litchi-drawingml/src/ink/codec.rs:906 | attr | checked (default) | FF | n/a | yes | `result.is_some()` guard | none | InkML part parse (15 callers) |
| 38 | crates/litchi-drawingml/src/ink/codec.rs:1040 | validate_element | checked (default) | FF; cap 256 | n/a | yes | pairwise `.any(expanded_attribute_names_equal)`: QUADRATIC but capped 256 | 2 `resolve_attribute` per pair | InkML part parse; first on every element (4 callers) |
| 39 | crates/litchi-drawingml/src/ink/codec.rs:1250 | has_namespace_declaration | checked (default) | SC (as #27; absent name, full scan) | OTHER (existence test) | no (bool) | none | up to ~4 calls per element via `is_namespace` | GUARDED: callers (`require_root`, `validate_unknown_prefix`, `is_namespace` in `frame_for_start`/`parse_empty`) run after #38 (FF, cap 256) |
| 40 | crates/litchi-drawingml/src/model3d/codec/xml.rs:545 | reference_plain | checked (default) | FF | n/a | yes | slot guard | per prefixed attr `is_relationship_prefix`: linear over merged namespaces (<= ~512) | model3d blip fragment re-parse (slice already admitted by the NsReader pass) |
| 41 | crates/litchi-drawingml/src/model3d/codec/xml.rs:604 | declarations_inner | checked (default) | FF | n/a | yes | `declarations.iter().any` on prefix: QUADRATIC in xmlns count, bounded 256 (NsReader cap in `read()`; plain re-parses only see NsReader-admitted tags) | `merge_namespaces` `iter_mut().find` per local decl (<=256x512) | model3d part parse (via `declarations` x2, `declarations_plain` x2) |
| 42 | crates/litchi-drawingml/src/svg_blip.rs:1548 | complete_raw_plan | checked (default) | FF; post-loop `declared.len() <= 256` check | n/a | yes | none | `declared.iter().any` per visible context binding (<=256 x <=16K) | contextual svgBlip write plan (retained fragment) |
| 43 | crates/litchi-drawingml/src/svg_blip.rs:1670 | parse_root | checked (default) | FF; cap 512 (after push) | n/a | yes | none | per-attr resolver lookups | svgBlip `codec::read` (NsReader) |
| 44 | crates/litchi-drawingml/src/svg_blip.rs:1797 | parse_contextual_root | checked (default) | FF; cap 512 (after push) | n/a | yes | none | per-attr ContextResolver lookups | svgBlip `read_contextual` (DOCX/XLSX drawings) |
| 45 | crates/litchi-drawingml/src/svg_blip.rs:1852 | contextual_reference | checked (default) | FF | n/a | yes | slot guard | none | same as #44 |
| 46 | crates/litchi-drawingml/src/svg_blip.rs:2173 | validate_contextual_attributes | checked (default) | FF | n/a | yes | none | per prefixed attr `scope.resolve` | contextual svgBlip children |
| 47 | crates/litchi-drawingml/src/svg_blip.rs:2411 | reference | checked (default) | FF (length pre-scan) | n/a | yes | none | then 2 more fail-fast passes via ooxml-common `relationships::attribute_value` | svgBlip `codec::read` |
| 48 | crates/litchi-drawingml/src/svg_blip.rs:2443 | declarations (helper) | checked (default) | FF | n/a | yes | `result.iter().any` on prefix: QUADRATIC in xmlns count. NsReader path bounded 256; plain-Reader contextual path (`ContextResolver::new`, `collect_contextual_children`) has only a post-hoc 256 check, so Theta(D^2) before refusal unless the host pre-capped | none | `read` and `read_contextual`; DOCX host pre-scans with NsReader (256 cap); XLSX host not verified; direct public `read_contextual` callers uncapped (<=16 MiB) |
| 49 | crates/litchi-drawingml/src/theme/codec.rs:1166 | ensure_supported_attributes | checked (default) | FF | n/a | yes | none | `allowed.contains` (constant) | theme part scheme elements (6 callers) |
| 50 | crates/litchi-drawingml/src/theme/family/codec.rs:166 | validate_extension_attributes | checked (default) | FF | n/a | yes | `uri_seen` guard | none | theme-family a:ext (after #52) |
| 51 | crates/litchi-drawingml/src/theme/family/codec.rs:479 | parse_root_attributes | checked (default) | FF | n/a | yes | `Option::replace` guards | none | theme-family root |
| 52 | crates/litchi-drawingml/src/theme/family/codec.rs:521 | validate_element_attributes | checked (default) | FF; cap 256 (count after `?`), xmlns cap 256 | n/a | yes | `seen.iter().any` (Vec): QUADRATIC, capped 256 | per-attr `resolve_attribute` | theme-family part, every element |
| 53 | crates/litchi-drawingml/src/theme/family/part.rs:564 | empty_container_after_removal | checked (default) | FF; `Ok(false)` at first non-xmlns/uri | n/a | yes | none | none | theme-family removal edit (re-parse of admitted source) |
| 54 | crates/litchi-drawingml/src/theme/family/part.rs:1524 | extension_uri | checked (default) | FF | n/a | yes | `uri.is_some()` guard | none | theme part a:ext classification (5 callers) |
| 55 | crates/litchi-drawingml/src/theme/family/part.rs:2252 | apply_namespace_declarations | checked (default) | FF | n/a | yes | `scope.iter().rev().take(added).any`: QUADRATIC in xmlns count, capped 256 (resolver cap set to 256) | none | theme part scan, every element |
| 56 | crates/litchi-drawingml/src/theme/family/part.rs:2304 | validate_element_attributes | checked (default) | FF; cap 256 | n/a | yes | `seen.iter().any` (Vec): QUADRATIC, capped 256 | `lookup_binding` linear per prefixed attr over scope (<= ~513) | theme part scan, every element |
| 57 | crates/litchi-drawingml/src/theme/family/part.rs:2379 | namespace_declaration_names | checked (default) | TOLERANT (`.flatten()`) | OTHER: collects all xmlns declaration names | no (returns Vec) | none | none | GUARDED: only in `classify_start`/`classify_empty`, which run after #56 (FF, cap 256) on the same element |
| 58 | crates/litchi-drawingml/src/transform/codec.rs:358 | validate_attributes | checked (explicit) | FF | n/a | yes | none | `allowed.contains` (constant) | a:xfrm parse (3 callers) |
| 59 | crates/litchi-formula/src/omml/handlers/accent.rs:18 | handle_start | checked (default) | TOLERANT: `.filter_map(a.ok()).collect()` | ALL: every Ok attr (duplicate Errs dropped, first occurrence kept); `get_attribute_value` first-wins; `parse_*_properties` plain assignment (last-wins only between `x`/`m:x` aliases) | no (returns `()`; caller in parser.rs returns Result) | none | re-iterates a tag parser.rs:157/595 already iterated (2x); constant number of linear lookups | OMML equation to LaTeX (`omml_to_latex`, markdown export of DOCX/PPTX equations); no size or attribute cap |
| 60 | crates/litchi-formula/src/omml/handlers/delim.rs:18 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 61 | crates/litchi-formula/src/omml/handlers/eq_arr.rs:20 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 62 | crates/litchi-formula/src/omml/handlers/fraction.rs:17 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 63 | crates/litchi-formula/src/omml/handlers/group_char.rs:19 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 64 | crates/litchi-formula/src/omml/handlers/matrix.rs:18 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 65 | crates/litchi-formula/src/omml/handlers/nary.rs:18 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 66 | crates/litchi-formula/src/omml/handlers/spacing.rs:18 | handle_start | checked (default) | TOLERANT (same) | ALL (same) | no | none | as #59 | as #59 |
| 67 | crates/litchi-formula/src/omml/parser.rs:157 | handle_start_element | checked (default) | TOLERANT (same) | ALL (same) | yes | none | then dispatches to a handler (#59-#66) that iterates the same tag again | every OMML start tag |
| 68 | crates/litchi-formula/src/omml/parser.rs:595 | handle_empty_element | checked (default) | TOLERANT (same) | ALL (same) | yes | none | `parse_attributes_batch`: ~50 linear lookups (O(50n)); plus handler re-iteration | every OMML empty tag |
| 69 | crates/litchi-ole-common/src/custom_xml/xml.rs:190 | required_attribute | checked (explicit) | FF | n/a | yes | `Option::replace` guard | per-attr `resolve_attribute` | legacy custom-XML data-store properties parse (2 callers) |
| 70 | crates/litchi-ole-common/src/custom_xml/xml.rs:223 | reject_other_attributes | checked (explicit) | FF | n/a | yes | none | per non-xmlns attr `resolve_attribute` | same (3 callers) |
| 71 | crates/litchi-ppt/src/slide_round_trip.rs:482 | parse_color_mapping_values | checked (explicit) | FF | n/a | yes (`Result<_, String>`) | `values[i].replace` guard | none | PPT binary round-trip color-mapping XML atom at open |
| 72 | crates/litchi-ppt/src/slide_round_trip.rs:758 | validate_xml_attributes | checked (explicit) | FF | n/a | yes | none | none | PPT round-trip XML, every element |
| 73 | crates/litchi-sign/src/xml.rs:1383 | start | checked (explicit) | FF; cap `max_attributes` (default 256) checked before the `?` on the pushed count | n/a | yes | expanded-name `HashSet` (std SipHash) after the loop | none in loop (per element clones parent namespace BTreeMap) | package XML-DSig signature part parse |
| 74 | crates/litchi-spreadsheet-drawing/src/chart/codec.rs:103 | user_shapes_ids | checked (default) | FF | n/a | yes | none (HashSet of r:ids) | per-attr `resolve_attribute` | chart userShapes part (after MCE `process_ooxml`) |
| 75 | crates/litchi-spreadsheet-drawing/src/chart/codec.rs:654 | fragment_ids | checked (default) | FF | n/a | yes | none | per-attr `resolve_attribute` | re-parse of chart-reader captured fragments (`as_xml`) |
| 76 | crates/litchi-spreadsheet-drawing/src/shape/reader.rs:885 | retain_unknown_attributes | checked (default) | FF | n/a | yes | none (Vec; count cap applied on push afterwards) | `known.iter().any` (constant) | XLSX drawing shape parse (2 method calls) |
| 77 | crates/litchi-spreadsheet-drawing/src/shape/reader.rs:1335 | any_truthy_attribute | checked (default) | FF (no early exit) | n/a | yes | none | none | XLSX drawing locks element |
| 78 | crates/litchi-spreadsheet-drawing/src/shape/tests.rs:283 | retains_unknown_objects_attributes_and_drawingml_children | NOT-QX (`Opaque::attributes() -> &[UnknownAttribute]`) | n/a | n/a | n/a | n/a | n/a | TEST |
| 79 | crates/litchi-spreadsheet-drawing/src/shape/tests.rs:284 | same | NOT-QX | n/a | n/a | n/a | n/a | n/a | TEST |
| 80 | crates/litchi-xldm/src/codec.rs:1328 | make_node | checked (explicit) | FF (counts non-xmlns attrs into the node; no cap) | n/a | yes | none | none | XLDM storage header/directory/partitions/backup-log XML (`inspect`) |
| 81 | crates/litchi-xldm/src/identity.rs:2031 | relationship_xml_frame | checked (default) | FF | n/a | yes | none (`class` plain assignment; dups unreachable) | none | XLDM OLAP relationship XML scan (2 callers) |
| 82 | crates/litchi-xldm/src/metadata/codec.rs:263 | make_xml_node | checked (explicit) | FF | n/a | yes | `attributes.iter().any(name == key)` (Vec): QUADRATIC and UNCAPPED; Theta(n^2) on n DISTINCT attributes (no duplicates needed); redundant with quick-xml's check | none | XLDM section 2.5 metadata `.tbl.xml` (`metadata::parse_file`); input <=16 MiB, node count capped, attributes per element not capped |
| 83 | crates/litchi-xldm/src/olap.rs:1108 | make_node | checked (explicit) | FF | n/a | yes | `attributes.iter().any(existing == key)` (Vec): QUADRATIC and UNCAPPED (as #82) | none | XLDM OLAP XML (`olap::parse_file`); input <=32 MiB, attributes per element not capped |
| 84 | crates/litchi-xlsb/src/cell_values/drawing_transfer.rs:1043 | inspect_element | checked (explicit) | FF | n/a | yes | none | per-attr `resolve_attribute`; plus fail-fast `unqualified_attribute` passes | drawing/chart transfer: source drawing XML inspection |
| 85 | crates/litchi-xlsb/src/cell_values/drawing_transfer.rs:1064 | namespace_declarations | checked (explicit) | FF | n/a | yes | BTreeMap insert (no refusal) | none | same |
| 86 | crates/litchi-xlsb/src/cell_values/drawing_transfer.rs:1350 | rewrite_element | checked (explicit) | FF | n/a | yes | BTreeSet `present`, O(log n) | per-attr resolve + BTreeMap lookup | same (rewrite) |
| 87 | crates/litchi-xlsb/src/cell_values/drawing_transfer.rs:1652 | xml_space | checked (default) | FF; returns at first `xml:space` | n/a | yes | none | none | chart XML compaction during transfer |
| 88 | crates/litchi-xlsb/src/comments/threaded/codec/mod.rs:718 | raw_attributes | checked (explicit) | FF | n/a | yes | none | `known.iter().any` (constant) | threaded comments / persons part (7 callers) |
| 89 | crates/litchi-xlsb/src/host/drawing.rs:772 | relationship_attribute | checked (default) | FF; returns at first r:id/embed/link | n/a | yes | none | per-attr `resolve_attribute` | XLSB host drawing part (2 callers) |
| 90 | crates/litchi-xlsb/src/timeline/codec.rs:68 | attributes | checked (explicit) | FF | n/a | yes | BTreeMap insert, refuses duplicate (redundant), O(log n) | none | timeline part parse |

## Helper functions that wrap attribute iteration (callers inherit the mode)

Call counts are call occurrences excluding the definition (grep over the crate `src`, or over the
defining file for private helpers).

| helper | mode | calls | notes |
|---|---|---|---|
| drawingml `chart/reader/codec/validation.rs:421` `get_attr` | SC (tolerant) | 50 direct | 24 inside validation.rs wrappers + 26 in chart_reader/{plot_area 21, document 2, series 2, legend 1}. The 24 wrappers each call it once and have ~148 callers in total: `parse_bool_attr` 55, `required_u32_attr` 25, `bounded_percentage_u32_attr` 10, `required_f64_attr` 10, `optional_bool_attr` 6, `parse_grouping` 6, `required_named_f64_attr` 6, `parse_number_format` 4, `required_enum_attr` 4, `optional_u32_attr` 3, `parse_time_unit` 3, `optional_i32_attr` 2, `parse_data_label_position` 2, `parse_tick_mark` 2, 10 others 1 each |
| drawingml `chart/reader/model.rs` `make_fragment_root_self_contained` | TOLERANT | 2 | via `capture_fragment` (51 calls) and `capture_empty_fragment` (51 calls) |
| drawingml `ink/actions_edit.rs` `payload_has_namespace_declaration` | SC (tolerant) | 3 | :1767 root, :2014 per element, :2115 per attribute (nested in #29) |
| drawingml `ink/actions.rs` `has_namespace_declaration` | SC (tolerant) | 3 | guarded by `validate_element` |
| drawingml `ink/codec.rs` `has_namespace_declaration` | SC (tolerant) | 3 | guarded; `is_namespace` wrapper has 8 calls |
| drawingml `theme/family/part.rs` `namespace_declaration_names` | TOLERANT | 2 | guarded |
| formula `omml/attributes.rs` `get_attribute_value*`, `AttributeCache`, `parse_attributes_batch`, `parse_*_properties` | operate on an already-collected `&[Attribute]` slice (no iterator) | 32 `get_attribute_value*` calls; 27 `parse_*_properties`/`parse_attributes_batch` calls | linear lookups; inherit the TOLERANT collection done at #59-#68 |
| crypto `labels.rs` `optional_attribute` / `required_attribute` | FF | 3 / 3 | `required_attribute` wraps `optional_attribute` |
| crypto `labels.rs` `reject_non_namespace_attributes` | FF | 2 | |
| crypto `ooxml/agile.rs` `exact_attributes` | FF | 7 | |
| ole-common `custom_xml/xml.rs` `required_attribute` / `reject_other_attributes` | FF | 2 / 3 | |
| drawingml `blip.rs` `read_embed` | FF | 1 (`find_first_embed`, test-only callers) | public |
| drawingml `formatcode2.rs` `validate_attributes` / `only_namespace_attributes` / `preflight_namespace_limits` / `attribute_projection` / `local_namespace_prefixes` / `validate_declaration` | FF | 3 / 2 / 3 / 1 / 1 / 1 | |
| drawingml `chart/reader/model.rs` `relationship_attribute_value` | FF | 1 | |
| drawingml `chart/model.rs` `validate_chart_xml_fragment` | FF | 1 (macro `from_xml`) | |
| drawingml `chart/style/codec.rs` `make_node` | FF | 2 | |
| drawingml `color/codec.rs` `attributes` | FF | 2 | |
| drawingml `diagram/data/codec.rs` `attributes` / `diagram/definition.rs` `attribute` | FF | 3 / 5 | |
| drawingml `ink/actions.rs` `validate_profile_attributes` / `optional_attr` / `optional_xml_id` / `required_attr` / `validate_element` | FF | 8 / 3 / 4 / 12 / 5 | |
| drawingml `ink/actions_edit.rs` `validate_payload_attributes_with_namespaces` / `validate_payload_element` / `check_profile_attributes` | FF | 1 / 2 / 2 | first has the nested SC re-scan |
| drawingml `ink/authoring.rs` `preflight_attributes` / `required_attribute` | FF | 1 / 7 | |
| drawingml `ink/codec.rs` `attr` / `validate_element` | FF | 15 / 4 | |
| drawingml `model3d/codec/xml.rs` `declarations_inner` / `reference_plain` | FF | 2 / 1 | |
| drawingml `svg_blip.rs` `declarations` / `reference` / `contextual_reference` / `validate_contextual_attributes` | FF | 4 / 1 / 1 / 2 | `declarations` has the uncapped-before-refusal quadratic on the plain-Reader path |
| drawingml `theme/codec.rs` `ensure_supported_attributes` | FF | 6 | |
| drawingml `theme/family/codec.rs` `validate_extension_attributes` / `parse_root_attributes` / `validate_element_attributes` | FF | 2 / 2 / 2 | |
| drawingml `theme/family/part.rs` `extension_uri` / `apply_namespace_declarations` / `validate_element_attributes` / `empty_container_after_removal` | FF | 5 / 2 / 2 / 2 | |
| drawingml `transform/codec.rs` `validate_attributes` | FF | 3 | |
| spreadsheet-drawing `shape/reader.rs` `retain_unknown_attributes` / `any_truthy_attribute` | FF | 2 / 1 | |
| xldm `metadata/codec.rs` `make_xml_node` / `olap.rs` `make_node` / `codec.rs` `make_node` | FF | 1 / 1 / 1 | first two have the uncapped quadratic own-dup check |
| xlsb `comments/threaded/codec/mod.rs` `raw_attributes` | FF | 7 | |
| xlsb `cell_values/drawing_transfer.rs` `unqualified_attribute` | FF (wraps ooxml-common) | 4 | |
| xlsb `host/drawing.rs` `relationship_attribute` / `timeline/codec.rs` `attributes` | FF | 2 / 1 | |
| out of scope but used here: ooxml-common `relationships::attribute_value`, `xml::unqualified_attribute_value` | FF (checked default, `map_err(..)?`) | used by svg_blip, model3d, transform, xlsb, spreadsheet-drawing | callers inherit FF |
| quick-xml `BytesStart::try_get_attribute` (validation.rs:182, :248) | UNCHECKED + FF (`a?`) | 2 | linear; not affected by the duplicate scan |

## Summary

Counts over the 88 quick-xml sites (all production; the 2 NOT-QX sites are in a test file):
- FAIL-FAST: 72 (checked 71, unchecked 1 = `ink/actions_edit.rs:4479`)
- TOLERANT: 12 (11 unguarded = 10 formula + chart/reader/model.rs:103; 1 guarded = theme/family/part.rs:2379)
- SHORT-CIRCUIT: 4 (2 unguarded, 2 guarded)
- COUNT (pure `.count()`-style): 0. Fail-fast sites that also carry a per-element cap: formatcode2 :417 (xmlns 64), :524 (64), :1271 (64, count-first, counts Err items), chart/style :662 (64), ink/actions :2074 (256), ink/actions_edit :2245 (256), ink/authoring :860 (256), ink/codec :1040 (256), svg_blip :1670/:1797 (512 after push), theme/family/codec :521 (256), theme/family/part :2304 (256), sign :1383 (256); agile :1590 has a document-wide 4096 counter that counts the Err item before the `?`.
- Checks: 24 `.with_checks(true)`, 63 default (checked), 1 `.with_checks(false)`.

TOLERANT production sites:
- crates/litchi-drawingml/src/chart/reader/model.rs:103: chart fragment roots, collects all names (unguarded)
- crates/litchi-formula/src/omml/handlers/accent.rs:18, delim.rs:18, eq_arr.rs:20, fraction.rs:17, group_char.rs:19, matrix.rs:18, nary.rs:18, spacing.rs:18: OMML handler, collects all Ok attrs (unguarded; tag already iterated once by parser.rs)
- crates/litchi-formula/src/omml/parser.rs:157, :595: every OMML tag, collects all (unguarded)
- crates/litchi-drawingml/src/theme/family/part.rs:2379: xmlns names; GUARDED by :2304

SHORT-CIRCUIT production sites:
- crates/litchi-drawingml/src/chart/reader/codec/validation.rs:422 `get_attr`: quadratic when name absent/last (typical for optional `val`); 50 direct, ~148 indirect callers (unguarded)
- crates/litchi-drawingml/src/ink/actions_edit.rs:1798 `payload_has_namespace_declaration`: absent name, full scan (unguarded; nested per attribute in :2059)
- crates/litchi-drawingml/src/ink/actions.rs:2325 and crates/litchi-drawingml/src/ink/codec.rs:1250 `has_namespace_declaration`: absent name, but GUARDED by `validate_element` (FF, cap 256)

QUADRATIC own-duplicate checks:
- UNCAPPED (Theta(n^2) on distinct names, no duplicates needed): crates/litchi-xldm/src/metadata/codec.rs:263 and crates/litchi-xldm/src/olap.rs:1108 (`Vec::iter().any`, redundant with quick-xml's own check; inputs up to 16/32 MiB).
- Capped only after the loop: crates/litchi-drawingml/src/svg_blip.rs:2443 `declarations` on the plain-Reader contextual path (post-hoc 256 check); bounded when the host pre-scans with an NsReader (DOCX does); direct `read_contextual` callers are not bounded.
- Bounded by a cap (<=64^2 or <=256^2): chart/style/codec.rs:662, theme/family/codec.rs:521, theme/family/part.rs:2252 and :2304, model3d/codec/xml.rs:604 (NsReader 256), ink/actions.rs:2074 and ink/codec.rs:1040 (pairwise expanded-name compare, 2 resolver scans per pair), color/codec.rs:396 (tiny allowed list).

Other super-linear patterns:
- crates/litchi-drawingml/src/ink/actions_edit.rs:2059: nested full re-scan (#28) per Unknown-prefix `iact:`/`inkml:` attribute. Theta(n^2) with no duplicates, Theta(D^2*R) with them; no attribute cap before pass 3. Input: caller-authored OpaquePayload, up to 16 MiB. Also :2014 runs one tolerant scan per element before any fail-fast pass.
- chart/reader/model.rs:103: `existing_names.iter().any` per root namespace decl, O(K*n) with K <= ~259.
- formula parser.rs:595: `parse_attributes_batch` does ~50 linear lookups per element, and every handled OMML tag is iterated twice (parser plus handler). Constant factors on top of the quadratic.
- `get_attr` is called several times per chart element, and each call pays the full Theta(D*R).
- Many NsReader sites call `resolve_attribute` per attribute. That is a reverse scan of in-scope bindings: at most 256 per element, times depth.

Cross-cutting reachability notes:
- The chart reader (`chart/reader/codec/xml.rs:29`) runs `mce::process_markup_compatibility` first. That pass iterates fail-fast and checked (`litchi-ooxml-common/src/mce/codec.rs:771`), but it returns the input unchanged when the part does not contain the MCE namespace URI. A crafted chart part that omits MCE therefore reaches `get_attr` and #16 with duplicate-laden tags. `ChartXmlReader::read_event_into` validates attributes only on the chartSpace root.
- Chart fragments captured by #16 are validated fail-fast afterwards by #13 (`from_xml`), so the tag is refused, but only after the tolerant cost has been paid.
