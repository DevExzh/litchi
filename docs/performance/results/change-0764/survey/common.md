# 0764 survey: quick-xml attribute iterators in litchi-ooxml-common, litchi-opc, xml-minifier, xml-minifier-macros

Read-only survey. **All line numbers and code descriptions are for committed HEAD `1d1044e3ac`**;
the cited lines were re-checked with `git show HEAD:`. While the survey ran, another agent was editing
the worktree (uncommitted). At the last check the edited files were: `mce/codec.rs`
(it now routes lookups through `resolver(scope, ..)`), `mce/mod.rs`, `mce/model.rs` (adds
`Limits::max_attributes_per_element`, default 1024), `custom_xml/codec.rs`, `web/codec/xml.rs`,
`xml-minifier/src/audit.rs`, `audit/namespaces.rs` and xml-minifier tests. Working-copy line numbers in
those files no longer match this table; for example, the `audit.rs` sites have moved 2515→2571,
3150→3251, 3365→3466.

## quick-xml 0.41.0 facts checked in the registry source

* `BytesStart::attributes()` starts with duplicate checks on (`events/attributes.rs:1045`,
  `check_duplicates: true`).
* `check_for_duplicates` (`attributes.rs:1135`) does a linear scan while fewer than 32 keys have been seen.
  From 32 keys on it uses `check_for_duplicates_hashed` (`:1160`), an **unkeyed**
  `DefaultHasher::new()` 64-bit hash stored in an `IdentityHasher` set. On a hash hit it runs
  `keys.iter().find(..)` to locate the earlier key.
  **A duplicate costs the index of that name's first occurrence**. A duplicate is never pushed onto
  `keys`, so the list stays at D distinct names. D distinct names followed by R copies of the
  *last* one therefore cost Θ(D·R). Copies of the *first* name cost O(1) each. A crafted 64-bit
  collision would also force the scan on a non-duplicate, because the hash is unkeyed; that case is
  not analysed here.
* After `Err(Duplicated)` the state becomes `SkipEqValue`, and iteration continues with the next attribute.
* `try_get_attribute` uses `with_checks(false)` and `?`. None of the surveyed crates calls it.
* `NsReader` calls `NamespaceResolver::push` on every Start/Empty event **before** the caller sees
  the event (`reader/ns_reader.rs:83/87`). `push` iterates with `with_checks(false)`, **breaks at the
  first malformed attribute**, and refuses more than `max_declarations_per_element` declarations
  (default 256, `name.rs:529`) with `TooManyDeclarations`. As a result, every element that an
  NsReader delivers carries at most 256 `xmlns*` attributes, counting only up to its first malformed
  attribute. The only explicit setters are `xml_splice.rs:931` and `source_backed.rs:1195`, and both
  set 256.
* `resolve_prefix` (`name.rs:951`) does a reverse linear scan over **all** bindings on the stack,
  shadowed re-declarations included, so a lookup costs O(B) with B ≤ 256 × depth.
  `NamespaceResolver::bindings()` (`name.rs:1179`) costs **O(B²)**, because each binding scans every
  later binding for an override. `Clone` copies `buffer` and `bindings`.
* std `HashSet`/`HashMap` below means `RandomState`, a keyed SipHash-1-3.

Legend: FF = FAIL-FAST; C = namespace-chain length (every ancestor declaration, shadowed ones included);
B = quick-xml resolver bindings in scope; k = attributes (or declarations) on one element.

## Site table (60 grep hits: 58 quick-xml sites + 2 NOT-QX; 5 of the quick-xml sites are TEST)

| file:line | fn | checks | error mode | consumption | can-refuse | own-duplicate-check | other super-linear (same loop) | reachability |
|---|---|---|---|---|---|---|---|---|
| crates/litchi-ooxml-common/src/binding_tracker.rs:295 | `BindingTracker::push_scanned` | unchecked | FF (`let Ok else break`, silent `Ok(())`; mirrors quick-xml push); in-loop cap 256 declarations, checked before each `add` | declarations only → `add` | yes | none (a repeated `xmlns:p` on one element is pushed twice, like quick-xml) | none (`record_binding` O(1); `contains_xmlns` prefilter O(tag)) | every Start/Empty element in the private tracker scanners: `xml.rs` OMML scan, docx paragraph text and transaction, pptx shape text, `mce/fragment.rs` |
| crates/litchi-ooxml-common/src/binding_tracker.rs:681 | `with_in_scope_namespaces` (pub via `private`) | unchecked | FF (`break` on Err, silent) | declarations → `declared: Vec` | **no** (returns `BytesStart`) | n/a (membership: `declared.contains` per binding) | **O(U·D)**: `declared.contains(&prefix)` for every in-scope binding (U = distinct prefixes in scope; D ≤ 256 when the element came from an NsReader) | 4 callers in litchi-xlsx (sheet_view, named_sheet_view, ignored_errors, conditional_formatting capture of an element into a standalone fragment) |
| crates/litchi-ooxml-common/src/binding_tracker.rs:849 | `tests::resolver_trace` | unchecked | TOLERANT (`.flatten()`) | ALL (collect resolved names) | no | none | quick-xml `resolve_attribute` O(B) per attribute | TEST |
| crates/litchi-ooxml-common/src/binding_tracker.rs:928 | `tests::tracker_trace_keeping` | unchecked | TOLERANT (`.flatten()`) | ALL | no | none | tracker lookup, bounded | TEST |
| crates/litchi-ooxml-common/src/custom/codec.rs:399 | `validate_root` | checked | FF (`map_err ?`); in-loop cap `MAX_ATTRIBUTES` = 32, counted **before** unwrap (the Err item counts) | only `xmlns*` allowed; first other attribute is refused | yes | none (quick-xml) | none (cap 32 < hashing threshold) | docProps/custom.xml decode (`custom/package.rs:29`), NsReader |
| crates/litchi-ooxml-common/src/custom/codec.rs:438 | `parse_property_attributes` | checked | FF; same 32 cap | NAMES name/pid/fmtid, **guarded** (`if x.is_none()`, else refuse) | yes | Option guards | none | every `property` element of custom.xml |
| crates/litchi-ooxml-common/src/custom/codec.rs:505 | `validate_value_attributes` | checked | FF; same 32 cap | only `xmlns*` allowed | yes | none | none | every vt value element of custom.xml |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:355 | `NamespaceState::enter` | checked (default) | **COUNT**: `.count()` is count-first and **counts Err items**, so it iterates past every `Duplicated`; cap `limits.attributes` (100 000) is compared only **after** the full count | COUNT | yes | — | **Θ(D·R)** before any cap; input ≤ 4 MiB (`MAX_PROPERTIES_XML_BYTES`) | every Start/Empty element of xlsx Custom Data Properties parts and their extension fragments (parse and validate paths; 4 call sites) |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:366 | `NamespaceState::enter` | checked | FF (`map_err ?`) | declarations → push binding | yes | **QUADRATIC**: `self.bindings[scope..].iter().any(prefix ==)` per declaration; k ≤ 256 (NsReader cap), so ≤ 32 K compares per element | namespace-bytes cap (4 MiB) checked during | same elements, right after 355 |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:588 | `extension_root_opening` | checked (default) | **COUNT**: count-first, counts Err items, cap checked after | COUNT (sizes `declared`) | yes | — | **Θ(D·R)** | extension (`extLst`) fragment root, reached from `parse_properties` (`:684`) and rewrite (`:901`) |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:599 | `extension_root_opening` | checked | FF | declarations → `declared` | yes | none | caller `decorate_extension_fragment:466` does `declared.iter().any` per inherited binding: O(I·D), I ≤ seeded bindings ≤ `limits.attributes`, D ≤ 256 | same |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:1655 | `element_attributes` | checked (default) | **COUNT**: count-first, counts Err items, cap after. Every caller runs `enter` on the same element first, and `enter` refuses duplicates at 366, so in practice this count only runs on duplicate-free tags | COUNT | yes | — | Θ(D·R) in isolation; linear only because of the call order | 4 call sites (`:1076`, `:1205`, `:1481`, `:1500`), each right after `enter` |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:1667 | `element_attributes` | checked (default) | **COUNT**: a second full `.count()` used for `seen.try_reserve`; counts Err items | COUNT | yes | — | as for 1655 | same |
| crates/litchi-ooxml-common/src/custom_data/codec.rs:1669 | `element_attributes` | checked | FF | ALL (namespace, name, value, range) | yes | std `HashSet<(String,String)>` of expanded names | `NamespaceState::lookup` per attribute: reverse **linear** scan over all bindings, shadowed and seeded ones included; B ≤ seeded (≤ 100 K) + 256 per element × depth ≤ 128 | same |
| crates/litchi-ooxml-common/src/custom_xml/codec.rs:420 | `inspect_element` | checked | FF | NAMES/validation (expanded-name uniqueness only) | yes | std `HashSet<(Vec,Vec)>` | quick-xml `resolve_attribute` O(B) per attribute; B ≤ 256 × `MAX_DEPTH` 256; **no per-element attribute cap** (part ≤ 16 MiB) | every element of customXml item parts (`:285`, `:299`) |
| crates/litchi-ooxml-common/src/custom_xml/codec.rs:460 | `resolve_props_element` | checked | FF | ALL → `attributes` Vec | yes | std `HashSet<(String,String)>` | as 420; caller `required_attribute` does a linear `.find` per required name | itemProps parts (`:530`, `:543`) |
| crates/litchi-ooxml-common/src/custom_xml/codec.rs:722 | `validate_declaration` | checked | FF | NAMES in an ordered state machine (version/encoding/standalone; anything else refused) | yes | state machine | none (≤ 4 items) | `<?xml …?>` of custom XML parts |
| crates/litchi-ooxml-common/src/mce/alternative/codec.rs:287 | `choice_requires` | checked | FF, returns at the first `Requires` | NAMES (first `Requires`) | yes | (`validate_attributes` already refused a duplicate `Requires`) | caller `parse_branch`: per Requires token one `format!` + `resolve_element` O(B) + `requirements.iter().any` → **O(T·(B+T))**, where T ≤ distinct namespaces in scope (B ≤ 256 × depth 256) | public `mce::alternative::read`; **no workspace callers** outside its tests |
| crates/litchi-ooxml-common/src/mce/alternative/codec.rs:348 | `validate_attributes` | checked | FF | policy check (NAMES) | yes | bool flag for `Requires` | `resolve_attribute` O(B) per attribute; called twice per element (`validate_element` plus a Root/Any pass) | same (7 in-file calls) |
| crates/litchi-ooxml-common/src/mce/codec.rs:771 | `start` | checked | FF (`map_err ?`); **no attribute-count cap** (input ≤ 256 MiB) | ALL → `raw` (re-serialized by `write_start`) | yes | none for attributes; declarations: `with_local` **QUADRATIC** `local[..i].iter().any` with k **unbounded** (see scan analysis) | per attribute 2–3 `Namespaces::get` chain walks O(C) (`expand_parts` at `:837`, `:1347`, `:1218`, `:1108`); `with_local` O(k²)+O(k·C); `for_each_hoisted`/`shadowed_before` O(H²)+`declares` O(H·attrs) in `write_start`; `preserves_attribute` linear pattern scan per ignorable attribute | MCE preprocessing of every OOXML part containing the MCE namespace URI (`process_part`/`process_ooxml`/`process_str`/`process_markup_compatibility`: ~165 call lines across docx/pptx/xlsx/drawingml/xlsb/spreadsheet-drawing) |
| crates/litchi-ooxml-common/src/mce/fragment.rs:359 | `root_insertion_point` | unchecked | FF (`break`, silent `Ok`) | declarations on the fragment root → `declared` | yes | none | **no count cap** (the fragment root is never pushed through a tracker); `reserve_exact(..,1)` per push reallocates on every push; caller `make_self_contained` filter is **O(U·D)** | `self_contained_fragment` / `make_self_contained`: docx `namespace.rs:331`, `paragraph/model.rs:274`; pptx `transition.rs:306`, `shape/model.rs:295`, `transition/reader.rs:425`; xlsx `chain/codec.rs:774` |
| crates/litchi-ooxml-common/src/mce/stream.rs:2554 | `Processor::parse_element` | unchecked | FF (`map_err ?`); cap `max_attributes_per_event` (4096; hard ceiling 1<<20) checked **during** (before push); bytes cap during | ALL → `preliminary` | yes | expanded names: std `HashSet<&Name>` (`validate_duplicate_attributes`); declarations: `with_local` **QUADRATIC**, k ≤ 4096, run **twice** per element (raw chain `:2594`, semantic chain `:1839`) | `expand_attribute` → `Namespaces::get` O(C) per attribute; `reserve_vec(..,1)` (= `try_reserve_exact(1)`) per push reallocates on every push: O(k²) bytes moved, also in `local_namespaces`; `semantic_attrs` runs `preserves_attribute` (linear pattern scan) up to 2× and `is_ignorable` O(depth) up to 4× per attribute | streaming MCE: xlsx `raw/strings.rs`, `raw/styles.rs`, `raw/worksheet/x14ac.rs` |
| crates/litchi-ooxml-common/src/mce/tests.rs:2495 | `resolved` | checked | FF (`.expect` panics) | ALL | no | none | resolver O(B) per attribute | TEST |
| crates/litchi-ooxml-common/src/mce/tests.rs:2576 | `fragment_is_self_contained` | checked | FF (`let Ok else return false`) | OTHER (resolvability) | no | none | resolver O(B) per attribute | TEST |
| crates/litchi-ooxml-common/src/mce/tests.rs:3160 | `resolves` | checked | FF (`return false`) | OTHER | no | none | resolver O(B) per attribute | TEST |
| crates/litchi-ooxml-common/src/properties/read.rs:388 | `keyword_lang` | checked | FF | NAMES `xml:lang`, **guarded** (a second one is refused); any other non-xmlns attribute refused | yes | Option guard | ≤ 2 resolver lookups before refusal | `cp:keywords` in docProps/core.xml (NsReader) |
| crates/litchi-ooxml-common/src/properties/read.rs:518 | `validate_attributes` | checked | FF | NAMES `xsi:type` **guarded** (`has_w3cdtf`); others refused | yes | bool guard | ≤ 2 resolver lookups before refusal | every core.xml element |
| crates/litchi-ooxml-common/src/relationships/codec.rs:33 | `attribute_value` (pub helper) | checked | FF (full pass, no early return) | NAMES: one local name in the r namespace, **guarded** (a duplicate semantic hit is refused) | yes | Option guard | `resolve_attribute` O(B) for every attribute whose local name matches, even when the prefix is foreign (`continue`); an Unknown prefix allocates a Vec | 9 direct production calls plus 4 thin wrappers with 28 callers (see helper list) |
| crates/litchi-ooxml-common/src/ribbon/codec.rs:192 | `validate_attributes` | checked | FF; in-loop `count_node` (document-wide node budget 262 144) before unwrap | validation (expanded uniqueness) | yes | std `HashSet<(ns, local)>` | `resolve_attribute` O(B) per attribute; B ≤ 256 × depth 128 | every element of customUI (Ribbon) XML (`:72`, `:116`) |
| crates/litchi-ooxml-common/src/ribbon/codec.rs:348 | `validate_declaration` | checked | FF | NAMES state machine | yes | state machine | none | Ribbon `<?xml …?>` |
| crates/litchi-ooxml-common/src/spreadsheet_xml_maps/codec.rs:705 | `optional_attr` (private helper) | checked | FF (full pass) | NAMES, **guarded** (a duplicate is refused) | yes | Option guard | each wrapper call is a full pass: `parse_map` makes ~10 passes per element | xl/xmlMaps.xml (after `mce::process_ooxml`); 7 direct calls + wrappers |
| crates/litchi-ooxml-common/src/spreadsheet_xml_maps/codec.rs:773 | `only_attrs` | checked | FF | whitelist check | yes | none | `allowed.contains` over a fixed ≤ 10-item list | 4 calls (MapInfo, Schema, Map, DataBinding) |
| crates/litchi-ooxml-common/src/spreadsheet_xml_maps/codec.rs:790 | `namespace_attributes` | checked | FF | declarations → Vec | yes | none | caller `merged_bindings` does `find` per local declaration: O(P·L), ≤ 256 × 256 | MapInfo root, Schema, DataBinding (4 calls) |
| crates/litchi-ooxml-common/src/spreadsheet_xml_maps/codec.rs:823 | `add_inherited_bindings` | **checked** (`with_checks(true)`) | **TOLERANT** (`.filter_map(Result::ok)`) | declarations → std `HashSet<Vec<u8>>` (membership for inherited bindings) | **no** (returns `BytesStart`; both callers return `Result`) | HashSet (membership, not a duplicate check) | **Θ(D·R)**: nothing validated this element's attributes earlier (NsReader push is unchecked), and a part without the MCE namespace skips MCE entirely; the Writer then **retains the duplicate attributes verbatim** in the opaque payload | first opaque child of `Schema`/`DataBinding` in xl/xmlMaps.xml (`begin_capture`, `capture_empty`); part ≤ 32 MiB |
| crates/litchi-ooxml-common/src/web/codec/xml.rs:320 | `used_namespace_prefixes` | checked | FF | prefixes → std `HashSet` | yes | HashSet | none | retained web-extension fragment re-parse (`self_contained_fragment_with_limits`) |
| crates/litchi-ooxml-common/src/web/codec/xml.rs:563 | `push_element` | checked | FF; `string_bytes` cap checked during | ALL (declarations → `HashMap`; others → `raw_attributes`) | yes | std `HashSet` (declared prefixes) + std `HashSet<(ns, local)>` | `NamespaceScope::get` walks a chain of per-layer `HashMap`s: O(depth) per lookup (fine). For captured `extLst` fragments, `effective_namespaces` walks every layer: O(all declarations, shadowed included) per capture | every element of Office web-extension XML (plain Reader) |
| crates/litchi-ooxml-common/src/web/model/custom_functions.rs:635 | `reject_no_attributes` | checked | FF | only `xmlns*` allowed | yes | none | none | web-extension custom-functions parts (4 calls) |
| crates/litchi-ooxml-common/src/web/model/custom_functions.rs:650 | `parse_contains` | checked | FF | NAMES `val` **guarded** | yes | Option guard | ≤ 1 resolver lookup before refusal | containsCustomFunctions |
| crates/litchi-ooxml-common/src/web/model/custom_functions.rs:681 | `parse_background` | checked | FF | NAMES state/runtimeId **guarded** | yes | Option guards | ≤ 2 resolver lookups before refusal | backgroundAppData |
| crates/litchi-ooxml-common/src/xml.rs:122 | `unqualified_attribute_value` (pub helper) | checked | FF (full pass) | NAMES (one unprefixed name) **guarded** | yes | Option guard | each call is a full pass, so N lookups on one element cost N·k | **213 production call sites** (xlsx 94, pptx 60, spreadsheet-drawing 25, drawingml 15, xlsb 15, docx 4) plus 1 test |
| crates/litchi-opc/src/content_type.rs:362 | `required_attributes` | checked | FF (`?`) | NAMES Extension/PartName + ContentType: plain assignment, which **would OVERWRITE** a yielded duplicate (unreachable: a raw duplicate is an Err) | yes | none (quick-xml) | none | [Content_Types].xml at package open (`pkgreader.rs:881`, `:972`) |
| crates/litchi-opc/src/package/content_types.rs:404 | `without_part_overrides` | checked | FF | NAMES PartName, plain assignment (would OVERWRITE; unreachable) | yes | none | none | override removal on an already-validated content-types token |
| crates/litchi-opc/src/package/relationships.rs:860 | `relationship_id` | checked | FF, returns at the first `Id` | NAMES (first `Id`) | yes | none | none | relationship edit scan of source .rels (`scan_relationship_source:775`, `:789`) |
| crates/litchi-opc/src/package/relationships.rs:965 | `OwnedRelationships::without_relationships` | checked | FF, `break` at the first selected Id | NAMES Id → std `HashSet` lookup | yes | none | none | relationship removal (edit path) |
| crates/litchi-opc/src/package/relationships.rs:1575 | `check_relationship_attributes` | checked | FF | OTHER (byte limits on every value) | yes | none | none | `relationship_xml_metrics` recount of ingress-validated .rels |
| crates/litchi-opc/src/part.rs:572 | `XmlPart::find_elements_with_attrs` (pub) | checked | FF (`attr?`) | ALL → std `HashMap` (insert would overwrite; unreachable) | yes | none | none; no size or count caps | public API, **0 callers** |
| crates/litchi-opc/src/pkgreader.rs:2147 | `inspect_relationship_element` | checked | FF | NAMES Id/Type/Target/TargetMode: plain assignment (would OVERWRITE; unreachable) | yes | none | none; per-value byte caps during | every .rels part at package open |
| crates/litchi-opc/src/source_backed.rs:736 | `inspect_relationship_append_root_attributes` | checked | FF; cap 4096 counted before unwrap | validation of every root attribute (opaque attributes kept) | yes | **QUADRATIC** `seen_keys: Vec` `.contains` + **QUADRATIC** `seen_expanded: Vec` `.iter().any`; k ≤ 4096 | 2 resolver lookups per attribute (`validate_source_attribute_key`, then `resolve_attribute`); depth 1, so B ≤ 258 | non-canonical .rels relationship append/topology splice on source-backed packages (`:1316`, `:1420`) |
| crates/litchi-opc/src/source_backed.rs:857 | `inspect_relationship_append_child` | checked | FF; cap 4096 | NAMES Id/Type/Target/TargetMode **guarded**; others refused apart from `xmlns*` | yes | **QUADRATIC** `seen_keys: Vec` `.contains`; effectively k ≤ 4 + 256 (resolver cap) | none | same (`:1342`, `:1454`) |
| crates/litchi-opc/src/source_backed/content_types_plan.rs:764 | `override_part_name` | checked | FF | NAMES PartName, plain assignment (would OVERWRITE; unreachable) | yes | none | none | source-backed content-types plan (publication/edit) |
| crates/litchi-opc/src/source_backed/content_types_plan.rs:785 | `default_extension` | checked | FF, returns at the first `Extension` | NAMES | yes | none | none | same |
| crates/litchi-opc/src/xml_splice.rs:1102 | `validate_source_declaration` | checked | FF; cap 3 counted before unwrap | NAMES state machine | yes | state machine | none | source XML declaration |
| crates/litchi-opc/src/xml_splice.rs:1177 | `validate_source_element` | checked | FF; cap `MAX_SOURCE_XML_ATTRIBUTES_PER_ELEMENT` 4096 counted before unwrap | validation (all attributes) | yes | **QUADRATIC** `seen_keys: Vec` `.contains` (redundant with the checked iterator) + **QUADRATIC** `seen_expanded: Vec` `.iter().any`; k ≤ 4096, about 8.4 M compares each per element | 2 quick-xml resolver lookups per attribute, O(B) each, B ≤ 256 × 256 ≈ 65.5 K | `validate_source_xml`: every element of source-preserved XML parts, **re-run on the whole part after every OwnedXmlPart edit** |
| crates/litchi-opc/src/xml_splice/owned.rs:241 | `OwnedXmlPart::update_attributes` | checked | FF | NAMES via `pending.remove` (std HashMap; consumed, so guarded) | yes | HashMap | none in the loop; `validate_source_xml` of the output afterwards (`:331`) | format-owner attribute edits |
| crates/litchi-opc/src/xml_splice/owned.rs:601 | `OwnedXmlPart::insert_unqualified_attribute` | checked | FF (a matching name → refusal) | NAMES | yes | none | full `validate_source_xml` of the output afterwards | same |
| crates/litchi-opc/src/xml_splice/owned.rs:694 | `OwnedXmlPart::replace_attributes` | checked | FF | OTHER (match value spans to edits) | yes | none | same | same |
| crates/xml-minifier-macros/src/lib.rs:600 | `minify_xml` | checked | FF (`?`) | NAMES `xml:space`: plain assignment (would OVERWRITE; unreachable) | yes | none | none | **compile time only**: proc macro over developer-authored XML literals |
| crates/xml-minifier/src/audit.rs:2515 | `inspect_attributes` | checked | FF; document-wide `limits.attributes` checked after unwrap | NAMES `xml:space` (would OVERWRITE; unreachable) | yes | none here (`check_start` checks expanded-name uniqueness in O(k) with keyed hashes) | none | `verify_source`/`verify_authored`/`verify_reader`: OPC publication and authored-fragment audits (4 calls) |
| crates/xml-minifier/src/audit.rs:3150 | `package::verify` | — | NOT-QX (`item.attributes()` is the `Report` usize getter) | — | — | — | — | — |
| crates/xml-minifier/src/audit.rs:3365 | `stream_tests::authored_reader_preserves_authored_whitespace_policy` | — | NOT-QX, TEST (`report.attributes()` getter) | — | — | — | — | — |

## Scans that grow with attacker-controlled counts (the five named files)

### crates/litchi-ooxml-common/src/mce/codec.rs (process_markup_compatibility)
1. **`Namespaces::with_local` duplicate check** (`:341-345`): `local[..index].iter().any(..)` is
   **O(k²)** for k declarations on one element. **There is no bound before or during the scan.** The
   only check, `bindings > max_namespace_bindings` (4096), runs **after** the loop, and `bindings` counts only
   prefixes new to the scope. The check is also redundant: the checked iterator at `:771` has already
   refused raw duplicate `xmlns:p` names. Example: one root carrying 10⁵ `xmlns:aN` declarations
   costs about 5·10⁹ string compares before the refusal. Input ≤ `max_input_bytes` (256 MiB).
2. **`with_local` → `self.get(prefix)` per declaration** (`:349`): O(k·C), with the same (absent) bound.
3. **`Namespaces::get`** (`:318-335`): walks every ancestor layer and reverse-scans each layer's
   `local` Vec, so a lookup costs O(C). C counts every ancestor declaration, **shadowed re-declarations included**.
   Bounds: `max_depth` 256 is checked **before** each element; per-element declarations are held to
   about 4096 only by the post-loop bindings check, so C ≤ about 256 × 4096 ≈ 1 M. `max_namespace_bindings` does
   **not** bound C, because re-declaring an in-scope prefix is never counted. Lookups per element: the element
   name, 2–3 per non-xmlns attribute (`:837` directive scan, `:1347` `write_start`, `:1218`
   AlternateContent validation, `:1108` ProcessContent unwrap), one per directive token (≤ 4096 per
   element, checked during), and **one per `Requires` token** (`:1030-1037`), which `max_directive_tokens`
   does **not** count and which has no cap. Worst case per element: attrs × C, or T_requires × C.
4. **`for_each_hoisted` + `shadowed_before`** (`:376-431`), called by `write_start` (`:1395`) whenever
   `inherited.hoists()`: O(H²), where H = declarations on dropped AlternateContent/Choice/Fallback/
   ProcessContent layers, plus `declares(raw, p)` O(attrs) per hoisted binding. This runs **again for every
   emitted child** of a dropped wrapper, so the total is children × H². H ≤ the C bound; no per-element bound.
5. **Directive patterns** (`pattern_directive_matches` `:508-524` → `matches_pattern` `:1280-1285`)
   use `HashSet::iter().any` (**linear**, not a hash lookup) on each directive layer. A query costs
   O(P), with P ≤ depth × `max_directive_tokens` (4096 per element, checked during token counting).
   Queries happen per element (`preserves_element`, `processes`) and per ignorable attribute
   (`preserves_attribute`, `:1364`).
6. `Ctx::is_ignorable` (`:484`) costs O(depth) keyed-HashSet probes per call. Bounded by `max_depth`; fine.

### crates/litchi-ooxml-common/src/mce/stream.rs
1. **`Namespaces::with_local`** (`:1566-1609`): the same O(k²) duplicate scan plus k × `get`. Here k ≤
   `max_attributes_per_event` (4096 by default; the `validate` hard ceiling is 1<<20), checked **before** the scan (during
   `parse_element`). The scan runs **twice per element**: the raw chain at `:2594` and the semantic chain at `:1839`.
   The bindings check and the `max_context_bytes` check both run **after** the loop. Because the iterator is unchecked, a duplicate `xmlns:p`
   is caught only by this O(k²) scan. Each element can cost about 2 × 8.4 M compares, which adds up to
   about input × 2k across the document.
2. **`Namespaces::get`** (`:1547-1564`): O(C). C ≤ 256 × 4096, and C is also bounded by `max_context_bytes` (16 MiB of
   cumulative declaration bytes along the chain, shadowed ones included; checked after each element).
   Lookups per element: the name, **every attribute** (`expand_attribute` in `parse_element`, ≤ 4096), and each
   directive or Requires token (≤ 4096, checked during).
3. **`Context::matches` → `matches_pattern`** (`:1654-1670`, `:3238-3243`): a linear scan over each
   layer's HashSet, O(P), with P ≤ depth × 4096 and bounded by `max_context_bytes` (checked in `apply_directives` both
   before and after). `semantic_attrs` (`:2776-2790`) calls it up to twice per ignorable attribute,
   so a single element costs O(attrs × P).
4. `Context::is_ignorable` costs O(depth) per call, up to 4× per attribute. Bounded; fine.
5. `reserve_vec(..,1)` is `try_reserve_exact(1)` on every push to `preliminary` (`:2570`) and to `local`
   (`:2826`, `:2843`), so the Vec reallocates on every push: up to O(k²) bytes moved, k ≤ 4096.

### crates/litchi-ooxml-common/src/mce/fragment.rs
1. `at_offset` (`:75`) re-walks the document up to the offset **on every call**. There is no memoization, so F
   fragments cut from one document cost O(F × document). F is bounded only by the caller.
2. `walked_tracker` (`:299`): tracker `push` holds each element to 256 declarations (checked during the push).
   `declaration_count() > max_namespace_bindings` (4096) is checked **after** each element's push, so B ≤ about
   4354. `max_depth` is checked **before** the push.
3. `from_tracker` → `for_each_in_scope` costs **O(B·U)** (`Vec::contains`). The result cap
   (`bindings.len() >= max_namespace_bindings`) fires during the walk, but the walk still visits all B bindings.
   `reserve_exact(..,1)` reallocates on every push.
4. `root_insertion_point` (site `:359`): the fragment root's declarations are collected with **no
   count cap**, and the Vec reallocates on every push.
5. `make_self_contained` (`:183-187`): the filter costs **O(U·D)**. U ≤ `max_namespace_bindings` when the scope came from
   `from_tracker`; U is **uncapped** through `from_declarations` (used by pptx `transition/reader.rs:425` and docx
   `paragraph/model.rs:268`). D is uncapped. `max_output_bytes` is checked only **after** the scan.
6. `InScopeNamespaces::namespace` (`:157`) is an O(U) linear find per lookup (docx `numbering/codec.rs`: 5 calls).

### crates/litchi-ooxml-common/src/binding_tracker.rs
1. `push_scanned` (`:290`): linear, with 256 declarations per element checked during. `pop` (`:369`): `rposition`
   plus truncation, amortized O(k).
2. Resolution (`innermost_prefixed`/`search_prefixed`) is **bounded**: 4 cache slots, then a list of ≤ 64 or a tail of ≤ 64,
   then a BTreeMap lookup in O(log U). The index grows once per binding (the 0754 hardening). OK.
3. **`for_each_in_scope`** (`:408-424`): `seen: Vec`, `.contains` per binding, so **O(B·U)**. B covers every
   binding, shadowed ones included, and the tracker has **no aggregate cap**: callers must check `declaration_count()`.
   mce/fragment does so after each push. The docx `namespace.rs` callers (×3) are caller-dependent.
4. **`in_scope_declarations`** (`:644-662`): quick-xml `bindings()` is **O(B²)**, and `.position` dedupe adds O(U²).
   B ≤ 256 × the caller's depth. No bound is enforced here. 6 call sites: xlsx ×4, one per captured
   element; pptx transition reader; docx paragraph model.
5. **`with_in_scope_namespaces`** (site `:681`): O(U·D) through `declared.contains`.

### crates/litchi-opc/src/xml_splice.rs
1. **`validate_source_element`** (site `:1177`): the two Vec-based seen lists cost 2 × O(k²/2), with k ≤ 4096 (the cap is checked
   **during** the loop, and the Err item counts), about 8.4 M compares each per element. The raw-name list is redundant with the
   checked iterator.
2. Per attribute, 2 quick-xml lookups (`validate_source_attribute_key:1288` and `:1241`); per element, a name lookup
   (`validate_source_qname:1315`) plus `read_resolved_event`'s own lookup; per End event, one more lookup. Each lookup costs O(B), with B ≤
   256 (`set_max_declarations_per_element`, enforced by NsReader **before** the event) × depth (≤
   min(`max_xml_depth` 256, 65534), checked **after** the push) ≈ 65.5 K. Per element the worst case is about
   4096 × 2 × 65.5 K ≈ 5.4·10⁸. Across the part, only `max_xml_events` (1 M) and part bytes (512 MiB) bound it.
3. Every OwnedXmlPart edit re-validates the whole output (`owned.rs:331`, `:362`, `:434`, `elements.rs:208`,
   `:310`, `sequence.rs:169`, `:230`), so callers that apply edits one at a time pay O(edits × part).

### NsReader / NamespaceResolver users in the four crates
* **Attribute resolution per attribute (O(B) each):** custom_xml `inspect_element`/`resolve_props_element`
  (no attribute cap); ribbon `validate_attributes`; mce/alternative `validate_attributes` (2 passes per
  element) plus Requires tokens; xml_splice `validate_source_element` (2× per attribute); source_backed root
  attributes (2×, B ≤ 258); `relationships::attribute_value` (matching local names only); properties and
  web custom_functions (≤ 2 per element before refusal).
* **Element resolution per event** through `read_resolved_event`: custom/codec, mce/alternative, properties, ribbon,
  web custom_functions, opc content_type.rs, pkgreader.rs, source_backed.rs, content_types_plan.rs,
  xml_splice.rs.
* **Resolver cloned on every event** (`reader.resolver().clone()` copies all in-scope bindings and URI bytes):
  **spreadsheet_xml_maps/codec.rs:312**, for every non-captured event. Comments and whitespace outside the capture have
  no event cap, and root URIs can total megabytes, so the cost is O(events × declared bytes).
  **web/package/transaction.rs:606**, for every event of both the before and after documents.
* `NamespaceResolver::bindings()` (O(B²)): binding_tracker `in_scope_declarations`.
* **NsReader used only as a tokenizer** (the cost is the linear push): custom_data (3 readers; its own `NamespaceState::lookup`
  is a **linear reverse scan over all bindings, shadowed and seeded ones included**, run once per attribute and per element);
  opc package/content_types.rs, package/relationships.rs, xml_splice/owned*.rs.
* xml-minifier: no NsReader. The audit uses its own keyed-hash resolver with O(1) lookups (audit/namespaces.rs).

## Helper functions that wrap attribute iteration (a caller inherits the helper's mode)

Workspace-wide counts come from `grep -rn crates/`, excluding the definition and comments.

| helper | mode | calls |
|---|---|---|
| `litchi_ooxml_common::xml::unqualified_attribute_value` (xml.rs:116) | checked FF, full pass | **213 production** (xlsx 94, pptx 60, spreadsheet-drawing 25, drawingml 15, xlsb 15, docx 4) + 1 test. At least 20 thin wrappers downstream (e.g. xlsx `required_string`/`optional_bool`, drawingml `required_attr`, pptx `parse_side`) |
| `litchi_ooxml_common::relationships::attribute_value` (relationships/codec.rs:26) | checked FF, full pass + O(B) resolve per matching local name | **9 production** (drawingml svg_blip ×2 and model3d ×2; pptx actions, namespace, embedded; spreadsheet-drawing ×1; xlsx raw/namespace ×1) + 2 test. Wrappers: pptx `relationship_attribute_value` (8 calls), pptx `relationship_value` (5), spreadsheet-drawing `relationship_attribute_value` (2), xlsx `raw::namespace::relationship_attribute_value` (13) |
| `relationships::attribute_id` | wraps `attribute_value` | 0 production, 1 test |
| `binding_tracker::with_in_scope_namespaces` | unchecked FF (break) + O(U·D) | 4 (xlsx) |
| `binding_tracker::in_scope_declarations` (resolver helper) | quick-xml `bindings()` O(B²) + O(U²) | 6 (xlsx 4, pptx 1, docx 1) |
| `BindingTracker::for_each_in_scope` (tracker helper) | O(B·U) | 4 (docx `namespace.rs` ×3, mce/fragment ×1) |
| `mce::self_contained_fragment` / `InScopeNamespaces::make_self_contained` | unchecked FF + O(U·D) | 4 + 2 (see the mce/fragment row) |
| `litchi_opc::XmlPart::find_elements_with_attrs` (part.rs:558) | checked FF | 0 |
| spreadsheet_xml_maps private: `optional_attr` ← `required_attr`/`parse_bool_attr`/`parse_u32_attr`/`required_bool_attr`/`required_u32_attr` | checked FF, full pass each | optional 7, required 5, parse_bool 2, parse_u32 2, required_bool 5, required_u32 2 (all in-file) |
| spreadsheet_xml_maps private: `only_attrs` / `namespace_attributes` / **`add_inherited_bindings`** | FF / FF / **TOLERANT** | 4 / 4 / 2 |
| custom_data private: `NamespaceState::enter` (COUNT + FF) / `element_attributes` (COUNT ×2 + FF) / `extension_root_opening` (COUNT + FF) | as the rows above | 4 / 4 / 1 (through `decorate_extension_fragment`, 2 calls) |
| other private FF helpers | FF | mce/alternative `validate_attributes` 7, `choice_requires` 1; custom_xml `inspect_element` 2, `resolve_props_element` 2; ribbon `validate_attributes` 2; properties `validate_attributes` 2, `keyword_lang` 3; web `reject_no_attributes` 4, `parse_contains` 2, `parse_background` 2; opc `required_attributes` 2, `relationship_id` 2, `check_relationship_attributes` 2, `override_part_name` 1, `default_extension` 1, `inspect_relationship_element` 2, append root/child 2/2, `validate_source_element` 2, `validate_source_declaration` 2; audit `inspect_attributes` 4 |

## Summary

* 60 grep hits: **58 quick-xml sites** (53 production, 5 TEST) and 2 NOT-QX (audit.rs:3150, 3365).
* Checks, production sites: 49 checked, 4 unchecked (binding_tracker:295, :681; mce/fragment:359; mce/stream:2554).
  TEST: 3 checked, 2 unchecked.
* Error mode, production sites: **FAIL-FAST 48, TOLERANT 1, COUNT 4, SHORT-CIRCUIT 0**.
  TEST: 3 FAIL-FAST, 2 TOLERANT (unchecked `.flatten()`, linear).
* **TOLERANT (production):** spreadsheet_xml_maps/codec.rs:823 `add_inherited_bindings`: filter_map(ok), retains duplicates; Θ(D·R).
* **COUNT (production):** custom_data/codec.rs:355 (`enter`, count-first, counts Errs; Θ(D·R)),
  :588 (`extension_root_opening`, same), :1655 and :1667 (`element_attributes`, two counts; each is Θ(D·R) in
  isolation but in practice runs only after `enter` has refused duplicates).
* **SHORT-CIRCUIT:** none. The early-return-on-match sites (alternative:287, relationships:860/965,
  content_types_plan:785, owned:601) all stop at the first Err, so they are FF.
* **QUADRATIC own checks:** mce/codec `with_local` O(k²) with **k unbounded**; mce/stream `with_local` O(k²),
  k ≤ 4096, run twice per element; xml_splice:1177 two Vec seen lists, k ≤ 4096; source_backed:736 two Vec
  lists, k ≤ 4096; source_backed:857 one Vec list (effectively ≤ 260); custom_data:366, k ≤ 256.
  Membership products: binding_tracker `for_each_in_scope` O(B·U), `in_scope_declarations` O(B²),
  `with_in_scope_namespaces` O(U·D); fragment `make_self_contained` O(U·D) with D uncapped; alternative
  Requires O(T²); custom_data `decorate_extension_fragment` O(I·D).
