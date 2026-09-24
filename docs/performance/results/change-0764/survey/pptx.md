# litchi-pptx: quick-xml attribute-iterator call-site survey (0764)

Worktree `/home/zhuhe/code/litchi-worktrees/0764-xml-attribute-dos-hardening` at `1d1044e3ac` (quick-xml 0.41.0).
Read-only survey: no build, no cargo.

## Method and scope

- I ran `grep -rn -E '\.(html_)?attributes\(\)'` over `crates/litchi-pptx/src`. It found **130 sites in 65 files**, including one site split across lines at `backgrounds/codec.rs:61`.
- Every site's receiver is a quick-xml `BytesStart` taken from a `Start` or `Empty` event, so there are **no NOT-QX sites**. The crate's own functions named `attributes(..)` are free helper functions that take a `&BytesStart`. None of them is a method on a crate type.
- **The crate never calls these:** `.html_attributes()`, `.with_checks(false)`, `try_get_attribute`, `attributes_raw`.
- **Check mode:** all 130 sites check for duplicates. 89 sites call `.with_checks(true)` explicitly (`expl` in the table). 41 sites use the default `.attributes()`, which also checks (`dflt` in the table).
- **Test code:** I treated a site as test code when it is inside a `#[cfg(test)] mod … {}` brace range or in a file declared with `#[cfg(test)] mod x;`. Exactly **1 site is TEST**: `notes/codec.rs:720`.
- I verified caller counts with a scoped regex that excludes `fn` definitions and test regions.

Legend:
- **flag**: a scalar duplicate guard (`Option::is_some()` or a `bool`), which is O(1).
- **HashSet/HashMap**: always `std::collections` with the default SipHash `RandomState`.
- **ns**: the loop calls quick-xml `NamespaceResolver::resolve_attribute` once per prefixed attribute. That call is a reverse linear scan of the bindings in scope. `NsReader` caps declarations at 256 per element, but the bindings accumulate with depth. The cost is O(n·B): bounded, but not constant.
- **No pre-pass shield:** `litchi_ooxml_common` MCE preprocessing (`process_markup_compatibility`, `process_ooxml`, `process_str`) has a fast path. It returns the input untouched when the document does not contain the MCE namespace URI. So an MCE pass never protects a later tolerant loop, because an attacker can simply leave the namespace out.

## Site table

| # | file:line | fn | checks | error mode | consumption | can-refuse | own-dup-check | other super-linear in loop | reachability |
|---|---|---|---|---|---|---|---|---|---|
| 1 | `crates/litchi-pptx/src/animations/codec/validation.rs:296` | `check_attribute_count` | chk dflt | FAIL-FAST (count cap 64; counts Ok items only, `Err`→return) | COUNT | yes | — | none | every element in the animation timing parsers (4 callers) |
| 2 | `crates/litchi-pptx/src/animations/codec/xml/parser/wire.rs:363` | `p14_attribute` | chk expl | FAIL-FAST | NAMES (one p14 local) | yes | flag | ns (matching local only) | p14 attributes in timing XML |
| 3 | `crates/litchi-pptx/src/animations/package.rs:389` | `inspect_start` | chk expl | **COUNT** (`.count() > MAX_ATTRIBUTES(64)`): counts first, counts `Err` items, iterates everything before comparing with the cap | COUNT | yes | — | none | every slide element in pub `parse_package_slide` / `Sequence::from_package_slide`. **SHIELDED:** `Sequence::parse_slide_xml` has already run `check_attribute_count` (fail-fast, ≤64, no dups) on every element of the same bytes |
| 4 | `crates/litchi-pptx/src/animations/package.rs:675` | `relationship_attribute` | chk expl | FAIL-FAST | NAMES | yes | flag | ns | graphicFrame host attributes (`parse_package_slide`) |
| 5 | `crates/litchi-pptx/src/backgrounds/codec.rs:61` | `SlideBackground::from_xml` | chk dflt | **SHORT-CIRCUIT** (`.flatten().find_map`, stops at the first `prst` with a UTF-8 value). Quadratic when `prst` is absent or comes after the duplicates | NAMES `prst`, first wins | yes | — | none | pub `SlideBackground::from_xml` on caller XML; no internal prod caller (tests, `crates/litchi` example, integration test); no pre-pass |
| 6 | `crates/litchi-pptx/src/backgrounds/codec.rs:101` | `SlideBackground::parse_color` | chk dflt | **SHORT-CIRCUIT** (`.flatten()` + `return` at the first `val`). Quadratic when `val` is absent or after the dups | NAMES `val`, first wins | yes | — | none | `srgbClr` under bg solidFill/gs/pattFill (from `from_xml`) |
| 7 | `crates/litchi-pptx/src/backgrounds/codec.rs:110` | `SlideBackground::parse_color` | chk dflt | **SHORT-CIRCUIT** (same as #6) | NAMES `val`, first wins | yes | — | none | `schemeClr`, same path |
| 8 | `crates/litchi-pptx/src/backgrounds/codec.rs:143` | `SlideBackground::parse_gradient` | chk dflt | **TOLERANT** (`.flatten()`) | NAMES `pos`, **OVERWRITE** (plain assignment; the only guard is numeric-parse success) | yes | — | none | `gs` stops in bg gradFill |
| 9 | `crates/litchi-pptx/src/backgrounds/codec.rs:193` | `SlideBackground::parse_gradient_descriptor` | chk dflt | **TOLERANT** (`.flatten()`) | NAMES `ang`, **OVERWRITE** (`*angle = Some(..)` on every valid match) | **no** (returns `()`; caller `parse_gradient` returns `Result`) | — | none | `a:lin` in bg gradFill |
| 10 | `crates/litchi-pptx/src/backgrounds/codec.rs:205` | `SlideBackground::parse_gradient_descriptor` | chk dflt | **TOLERANT** (`.flatten()`) | NAMES `path`, **OVERWRITE** | **no** | — | none | `a:path` in bg gradFill |
| 11 | `crates/litchi-pptx/src/change_tracking/codec.rs:570` | `parse_id` | chk expl | FAIL-FAST | NAMES `val` (refuses others) | yes | flag | none | MS-PPTX 2.2.9 change-tracking id elements (`classify`) |
| 12 | `crates/litchi-pptx/src/chart/extension/codec/xml.rs:520` | `attributes` (helper) | chk expl | FAIL-FAST (cap 64 non-xmlns, checked before the scan) | ALL | yes | Vec `.iter().any` on (ns, name): QUADRATIC but **capped at 64** | ns | chartex part parse (3 callers) |
| 13 | `crates/litchi-pptx/src/comments/codec.rs:434` | `make_node` | chk expl | FAIL-FAST | ALL | yes | HashSet\<String\> | none | legacy `commentAuthors` / `comments` parts |
| 14 | `crates/litchi-pptx/src/comments/collaboration/codec.rs:998` | `make_node` | chk expl | FAIL-FAST | ALL (+ source spans) | yes | HashSet\<String\> | **QUADRATIC, no duplicates needed:** every non-xmlns attribute calls `attribute_value_span` and `attribute_full_span`. Each rescans the raw tag from its start (`attribute_ranges`), so a tag with n attributes costs O(n·L) bytes. No attribute cap, only the 8 MiB part cap. Also ns | legacy comment collaboration `locate()` (presence/threading transactions); runs on raw source before MCE |
| 15 | `crates/litchi-pptx/src/font/codec.rs:365` | `parse_face` | chk expl | FAIL-FAST | NAMES `r:id` (refuses others) | yes | flag | ns | `embeddedFont` face elements |
| 16 | `crates/litchi-pptx/src/font/codec.rs:402` | `collect_unqualified_attributes` | chk expl | FAIL-FAST | ALL (allow-list) | yes | HashMap insert | ns; `allowed.contains` over a constant slice | `embeddedFont` elements (2 callers) |
| 17 | `crates/litchi-pptx/src/font/package.rs:564` | `embedding_enabled` | chk expl | FAIL-FAST (returns at the first match) | NAMES | yes | — | none | presentation root `embedTrueTypeFonts` |
| 18 | `crates/litchi-pptx/src/master_layout/codec.rs:687` | `element_relationship_id` | chk expl | FAIL-FAST (returns the first `*:id`) | NAMES | yes | — | none | master/layout ID-list entries (2 callers) |
| 19 | `crates/litchi-pptx/src/master_layout/codec.rs:714` | `next_shape_id` | chk expl | FAIL-FAST | NAMES `id` (max) | yes | — | none | every `cNvPr` in master/layout XML (placeholder insert) |
| 20 | `crates/litchi-pptx/src/master_layout/package.rs:599` | `push_layout_id_entry` | chk expl | FAIL-FAST | NAMES `id` + `*:id` (plain assignment) | yes | — | outside the attribute loop: two linear `entries.iter().any` per entry, O(E²) over `sldLayoutId` entries (E ≤ 100 000 nodes) | slide master `sldLayoutIdLst` |
| 21 | `crates/litchi-pptx/src/master_layout/package.rs:718` | `push_master_id_entry` | chk expl | FAIL-FAST | NAMES `id` + `*:id` | yes | — | outside the attribute loop: O(E²) entry scans, as #20 | presentation `sldMasterIdLst` |
| 22 | `crates/litchi-pptx/src/master_layout/placeholder.rs:1981` | `validate_extension_uri` | chk expl | FAIL-FAST | NAMES `uri` (refuses others) | yes | flag | none (one earlier full pass via `unqualified_attribute_value`) | `p:ext` in the placeholder owner |
| 23 | `crates/litchi-pptx/src/media_parts/codec.rs:577` | `make_node` | chk expl | FAIL-FAST | ALL | yes | **Vec `.iter().any` on expanded (ns, name): QUADRATIC, no count cap.** Bounded only by the 4 MiB `MAX_STRING_BYTES` budget and 32 MiB XML | ns | slide media `load`/`store` (`document_conformance` parses the whole slide) |
| 24 | `crates/litchi-pptx/src/modern_comments/codec.rs:322` | `authors::known_attributes` | chk expl | FAIL-FAST | ALL (allow-list) | yes | HashMap insert | none | modern comment authors part |
| 25 | `crates/litchi-pptx/src/modern_comments/codec.rs:344` | `authors::validate_any_attributes` | chk expl | FAIL-FAST | OTHER (value bound) | yes | — | none | same |
| 26 | `crates/litchi-pptx/src/modern_comments/codec.rs:355` | `authors::no_non_namespace_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | none | same |
| 27 | `crates/litchi-pptx/src/modern_comments/codec.rs:372` | `authors::namespace_declarations_from` | chk expl | FAIL-FAST | ALL xmlns | yes | HashSet afterwards (`validate_namespaces`) | none | same |
| 28 | `crates/litchi-pptx/src/modern_comments/codec/comments/parser.rs:572` | `known_attributes` | chk expl | FAIL-FAST | ALL (allow-list) | yes | HashMap insert | none | modern comments part |
| 29 | `crates/litchi-pptx/src/modern_comments/codec/comments/parser.rs:594` | `validate_any_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | none | same |
| 30 | `crates/litchi-pptx/src/modern_comments/codec/comments/parser.rs:605` | `no_non_namespace_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | none | same |
| 31 | `crates/litchi-pptx/src/modern_comments/codec/comments/parser.rs:620` | `namespace_declarations_from` | chk expl | FAIL-FAST | ALL xmlns | yes | HashSet afterwards | none | same |
| 32 | `crates/litchi-pptx/src/modern_comments/wire/xml.rs:224` | `namespace_declarations` | chk expl | FAIL-FAST | ALL xmlns | yes | — | none | modern comments wire parse |
| 33 | `crates/litchi-pptx/src/modern_comments/wire/xml.rs:253` | `attributes` (helper) | chk expl | FAIL-FAST | ALL | yes | — (later `attribute()` lookups detect repeats with one linear scan per lookup; the number of lookups K is fixed) | none | same (2 callers) |
| 34 | `crates/litchi-pptx/src/notes/codec.rs:473` | `inspect_element` | chk expl | FAIL-FAST (document-wide cap `MAX_ATTRIBUTES` = 500 000 on non-xmlns Ok items) | ALL r: values | yes | — | ns | notes / notesMaster / theme parts |
| 35 | `crates/litchi-pptx/src/notes/codec.rs:720` | `inspect_element_oracle` | chk expl | FAIL-FAST, **TEST** | ALL | yes | — | ns | test oracle only |
| 36 | `crates/litchi-pptx/src/opened/copy_plan.rs:1006` | `validate_xml_surface` | chk dflt | FAIL-FAST | OTHER (namespace policy) | yes | — | ns | opened-presentation slide copy plan |
| 37 | `crates/litchi-pptx/src/opened/copy_plan.rs:1109` | `reject_mce` | chk dflt | FAIL-FAST | OTHER | yes | — | ns | copy/remove/cross-copy plans (7 callers) |
| 38 | `crates/litchi-pptx/src/opened/remove_plan.rs:722` | `validate_presentation_owner` | chk dflt | FAIL-FAST | OTHER | yes | — | ns | slide removal plan (presentation.xml) |
| 39 | `crates/litchi-pptx/src/opened/transaction.rs:1748` | `shape_relationship_ids` | chk dflt | FAIL-FAST | ALL r: values → BTreeSet | yes | — | ns (HashMap `rels().get`) | opened-presentation shape transfer fragments |
| 40 | `crates/litchi-pptx/src/opened/xml.rs:172` | `element_preserves_space` | chk dflt | FAIL-FAST | NAMES `xml:space` | yes | — | none | `compact_changed_slide_xml` |
| 41 | `crates/litchi-pptx/src/opened/xml.rs:201` | `write_compact_start` | chk dflt | FAIL-FAST | ALL (re-serialize) | yes | — | none | same |
| 42 | `crates/litchi-pptx/src/opened/xml.rs:1148` | `connector_connection_ids` | chk dflt | FAIL-FAST | NAMES `id` | yes | flag | none | stCxn/endCxn in transferred shapes |
| 43 | `crates/litchi-pptx/src/opened/xml.rs:1203` | `write_remapped_shape_start` | chk dflt | FAIL-FAST | ALL (re-serialize) | yes | — (a linear `shape_identity_from_attributes` pass once) | ns (in the write loop) | `remap_shape_fragment` |
| 44 | `crates/litchi-pptx/src/opened/xml.rs:1562` | `parse_slide_id` | chk dflt | FAIL-FAST | NAMES `*:id` | yes | flag | ns (id locals only); 2 earlier fail-fast helper passes | `slide_id_elements` (presentation.xml) |
| 45 | `crates/litchi-pptx/src/parts/presentation.rs:402` | `reject_slide_id_list_attributes` | chk dflt | FAIL-FAST | OTHER | yes | — | none | presentation.xml `sldIdLst` (open) |
| 46 | `crates/litchi-pptx/src/parts/slide.rs:251` | `validate_semantic_attributes` | chk expl | FAIL-FAST (byte cap) | OTHER | yes | — | none | semantic slide XML scan (`observe`, `validate_attributes`) |
| 47 | `crates/litchi-pptx/src/parts/slide.rs:280` | `validate_semantic_attribute_names` | chk expl | FAIL-FAST | OTHER | yes | — | ns | same |
| 48 | `crates/litchi-pptx/src/presentation/embedded/content_parts/codec.rs:863` | `relationship_value_span` | chk dflt | FAIL-FAST | NAMES `r:id` | yes | flag | ns (id only); one `attribute_span` after the loop | contentPart anchors |
| 49 | `crates/litchi-pptx/src/presentation/embedded/content_parts/codec.rs:994` | `black_white_mode` | chk dflt | FAIL-FAST | NAMES `p14:bwMode` | yes | flag | ns | contentPart elements |
| 50 | `crates/litchi-pptx/src/presentation/embedded/content_parts/codec.rs:1021` | `validate_attributes` | chk dflt | FAIL-FAST (count cap 64; Ok items) | COUNT | yes | — | none | content-part scans (6 callers) |
| 51 | `crates/litchi-pptx/src/presentation/embedded/controls/codec.rs:174` | `parse_control` | chk expl | **COUNT** (`.count() > MAX_XML_ATTRIBUTES(64)`): counts first, counts `Err` items, iterates everything before the cap | COUNT | yes | — | none | `p:control` under `p:controls`, pub embedded controls `load_slide` → `scan`. **Exposed:** only the MCE fast path runs before it; the fail-fast helper passes that follow run only after the full count |
| 52 | `crates/litchi-pptx/src/presentation/embedded/controls/codec.rs:345` | `attribute` (helper) | chk dflt | FAIL-FAST | NAMES (ns-qualified) | yes | flag | ns (matching local) | ActiveX descriptor root (6 calls) |
| 53 | `crates/litchi-pptx/src/presentation/embedded/controls/slide/codec.rs:515` | `find_attribute` | chk expl | FAIL-FAST | NAMES (+ span) | yes | flag | ns (matching local); `attribute_span` at most once | control slide transactions (9 calls) |
| 54 | `crates/litchi-pptx/src/presentation/embedded/ink_actions/codec.rs:678` | `requires_value` | chk dflt | FAIL-FAST | NAMES `Requires` | yes | flag | one `resolve_prefix` (O(B)) per Requires token | ink-action `mc:Choice` |
| 55 | `crates/litchi-pptx/src/presentation/embedded/ink_actions/codec.rs:741` | `capture_relationship_attributes` | chk expl | FAIL-FAST | ALL `*:id` | yes | — (checked later by `relationship_id`) | ns (id only) | ink-action contentPart |
| 56 | `crates/litchi-pptx/src/presentation/embedded/ink_actions/codec.rs:824` | `validate_attributes` | chk expl | FAIL-FAST (count cap 256; Ok items) | COUNT + OTHER | yes | two Vec `.iter().any` scans (prefixes, expanded names): QUADRATIC but **capped at 256** | ns | ink-action owner XML (2 callers) |
| 57 | `crates/litchi-pptx/src/presentation/embedded/ole/codec.rs:105` | `make_node` | chk expl | **COUNT** (`.count() > MAX_XML_ATTRIBUTES(64)`): counts first, counts `Err` items, iterates everything | COUNT | yes | — | none | **every element** of the slide in pub embedded OLE `load_slide` → `parse_tree`. **Exposed:** only the MCE fast path runs first; the fail-fast loop #58 runs after the count |
| 58 | `crates/litchi-pptx/src/presentation/embedded/ole/codec.rs:114` | `make_node` | chk dflt | FAIL-FAST | ALL | yes | — | ns | same |
| 59 | `crates/litchi-pptx/src/presentation/embedded/ole/slide/codec.rs:167` | `make_node` | chk expl | **COUNT** (`.count() > 64`): counts first, counts `Err` items, iterates everything | COUNT | yes | none (the raw `scan_attributes` that follows has no duplicate check) | ns inside the raw scanner | every element in OLE slide transaction `locate()`. **SHIELDED in production:** `Snapshot::load` has already run `load_slide`, so #57/#58 validated the same bytes; `build_after`'s second `locate` runs on self-generated XML |
| 60 | `crates/litchi-pptx/src/presentation/order.rs:835` | `has_mce_attribute` | chk dflt | FAIL-FAST | OTHER | yes | — | ns | slide-order metadata/presentation scan |
| 61 | `crates/litchi-pptx/src/presentation/order.rs:858` | `validate_root_attributes` | chk dflt | FAIL-FAST | OTHER (allow-list) | yes | — | none | presentation root (slide order) |
| 62 | `crates/litchi-pptx/src/presentation/source.rs:4024` | `validate_blip_extension_attributes` | chk expl | FAIL-FAST | NAMES `uri` | yes | flag | none | source-backed picture descriptor |
| 63 | `crates/litchi-pptx/src/presentation/source.rs:4109` | `namespace_declarations` | chk expl | FAIL-FAST | ALL xmlns | yes | — | none | same (`p:pic` root) |
| 64 | `crates/litchi-pptx/src/presentation/source.rs:4192` | `validate_blip_attributes` | chk expl | FAIL-FAST | OTHER | yes | flag (`cstate`) | ns | same |
| 65 | `crates/litchi-pptx/src/presentation/source.rs:4311` | `is_opaque_drawing_extension` | chk expl | FAIL-FAST | NAMES `uri` | yes | flag | none | full-slide picture relationship scan |
| 66 | `crates/litchi-pptx/src/presentation/source.rs:4537` | `validate_full_slide_blip_attributes` | chk expl | FAIL-FAST | NAMES `r:embed`/`r:link` | yes | flags | ns | same |
| 67 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle.rs:1254` | `ext_list_contains_only_svg` | chk expl | **SHORT-CIRCUIT** (`.next().is_some()`: first item only, O(1), cannot reach a duplicate) | OTHER (presence test) | yes | — | none | SVG detach (`extLst` opening tag) |
| 68 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle.rs:1604` | `relationship_id_is_referenced_elsewhere` | chk expl | FAIL-FAST | ALL values (substring test) | yes | — | none | SVG attach/detach topology scan |
| 69 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs:146` | `Namespaces::preflight` | chk expl | FAIL-FAST (count cap 256; Ok items) | COUNT | yes | — | none | SVG owner `locate_all` (every element) |
| 70 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs:224` | `Namespaces::push` | chk expl | FAIL-FAST | ALL xmlns | yes | HashMap prefix stack | none | same |
| 71 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs:413` | `root_fragment_namespace_info` | chk expl | FAIL-FAST | ALL xmlns | yes | — | none | picture fragment root |
| 72 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs:1140` | `namespace_complete_element_fragment` | chk expl | FAIL-FAST | ALL xmlns | yes | — | after the loop, inherited × declared `.iter().any`; bounded by the upstream caps (≤16 384 × 256) | picture/SVG fragments |
| 73 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs:1336` | `extension_uri_is_supported` | chk expl | FAIL-FAST | NAMES `uri` | yes | flag | none | `a:ext` classification |
| 74 | `crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs:1372` | `validate_start_attributes` | chk expl | FAIL-FAST | OTHER | yes | Vec `.iter().any` on expanded names: QUADRATIC but **capped at 256** by `preflight` | none (resolver is a HashMap) | SVG owner `locate_all` (every element) |
| 75 | `crates/litchi-pptx/src/presentation/source_cross_copy.rs:2663` | `direct_embedded_images` | chk dflt | FAIL-FAST | NAMES `uri` | yes | flag | ns | cross-package slide copy (graphicData) |
| 76 | `crates/litchi-pptx/src/presentation/source_cross_copy.rs:2723` | `direct_embedded_images` | chk dflt | FAIL-FAST | NAMES `r:id` | yes | flag | ns | same (`c:chart`) |
| 77 | `crates/litchi-pptx/src/presentation/source_cross_copy.rs:2881` | `direct_embedded_images` | chk dflt | FAIL-FAST | NAMES `r:embed` | yes | flag | ns | same (`a:blip`) |
| 78 | `crates/litchi-pptx/src/presentation/source_cross_copy.rs:5067` | `validate_xml_with_policy` | chk dflt | FAIL-FAST | NAMES `val` | yes | flag | none | same (`p14:creationId`) |
| 79 | `crates/litchi-pptx/src/presentation/source_cross_copy.rs:5152` | `validate_xml_with_policy` | chk dflt | FAIL-FAST | OTHER (policy) | yes | — | ns | same (every element) |
| 80 | `crates/litchi-pptx/src/presentation/transition.rs:70` | `has_mce_markup` | chk expl | FAIL-FAST | OTHER | yes | — | ns | source-backed transition `validate_source` |
| 81 | `crates/litchi-pptx/src/presentation/transition.rs:662` | `root_has_namespace` | chk expl | FAIL-FAST (returns at the first match) | NAMES `xmlns:p` | yes | — | none | AlternateContent root |
| 82 | `crates/litchi-pptx/src/presentation/transition.rs:1752` | `validate_transition_element` | chk dflt | FAIL-FAST | NAMES (allow-list) | yes | Vec `seen.iter().any`, bounded by the allow-list size (≤ a handful) | ns (p14 `dur` only) | transition subtree validation |
| 83 | `crates/litchi-pptx/src/presentation_properties/codec.rs:192` | `make_node` | chk expl | FAIL-FAST | ALL | yes | — | **SUPER-LINEAR, uncapped when the MCE namespace is absent:** (a) the parent's `bindings.clone()` per node, O(B) time and retained memory per node, so O(nodes·B) overall; (b) `bindings.iter_mut().find` per xmlns, O(k·B), which is k² for k declarations on one tag; (c) a reverse-linear `resolve()` per prefixed attribute, O(m·B) | `presProps.xml` `Properties::parse` (plain `Reader`, 8 MiB, 100 000 nodes) |
| 84 | `crates/litchi-pptx/src/presentation_properties/math/codec.rs:261` | `required_value` | chk expl | FAIL-FAST | NAMES `m:val` | yes | flag | ns | presProps math ext |
| 85 | `crates/litchi-pptx/src/presentation_properties/math/codec.rs:289` | `no_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | ns (once) | same |
| 86 | `crates/litchi-pptx/src/presentation_properties/metadata/changes/codec.rs:866` | `root_namespaces` | chk expl | FAIL-FAST | ALL xmlns | yes | HashSet | none | changesInfo part |
| 87 | `crates/litchi-pptx/src/presentation_properties/metadata/changes/codec.rs:895` | `known_attributes` | chk expl | FAIL-FAST | ALL (allow-list) | yes | HashSet | `known.contains` over a constant slice | same (5 callers) |
| 88 | `crates/litchi-pptx/src/presentation_properties/metadata/changes/codec.rs:933` | `any_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | none | same |
| 89 | `crates/litchi-pptx/src/presentation_properties/metadata/custom_show/codec.rs:31` | `List::parse_xml` | chk dflt | **TOLERANT** (`.flatten()`) | NAMES `name`, `id`: **OVERWRITE** (plain assignment; a bad `id` falls back to 0) | yes | — | none | pub `custom_show::List::parse_xml` on caller XML; **0 internal prod callers**; the `mce::process_str` pre-pass runs only when the MCE namespace is present |
| 90 | `crates/litchi-pptx/src/presentation_properties/metadata/custom_show/codec.rs:51` | `List::parse_xml` | chk dflt | **TOLERANT** (`.flatten()`) | OTHER: appends one slide per matching `r:id`/`id` attribute (accumulates, does not overwrite) | yes | — | none | same (`sld` under `custShow`) |
| 91 | `crates/litchi-pptx/src/presentation_properties/metadata/custom_show/wire.rs:747` | `attributes` (helper) | chk expl | FAIL-FAST | ALL | yes | HashSet | none | custom-show wire parse (2 callers) |
| 92 | `crates/litchi-pptx/src/presentation_properties/metadata/designer_tags/codec.rs:500` | `attributes` (helper) | chk expl | FAIL-FAST | NAMES + ALL xmlns | yes | flags (`id`/`uri` `replace().is_some()`) | one fail-fast `relationship_attribute_value` pass after the loop | Designer-tag owner XML |
| 93 | `crates/litchi-pptx/src/presentation_properties/metadata/guides/codec.rs:354` | `element_uri` | chk expl | FAIL-FAST | NAMES `uri` | yes | flag | none | presentation guides ext |
| 94 | `crates/litchi-pptx/src/presentation_properties/metadata/guides/codec.rs:607` | `make_node` | chk expl | FAIL-FAST | ALL | yes | — | **SUPER-LINEAR (same DOM pattern as #83)** | `Guides::from_xml` / `rewrite_source` (plain `Reader`, 8 MiB) |
| 95 | `crates/litchi-pptx/src/presentation_properties/metadata/handout/codec.rs:28` | `Master::parse_xml` | chk dflt | **TOLERANT** (`.flatten()`) | NAMES `hdr`/`ftr`/`sldNum`/`dt`: **OVERWRITE** | yes | — | none | pub `handout::Master::parse_xml`; **0 internal prod callers** (examples, integration test); MCE pre-pass only when the namespace is present |
| 96 | `crates/litchi-pptx/src/presentation_properties/metadata/handout/codec.rs:49` | `Master::parse_xml` | chk dflt | **TOLERANT** (`.flatten()`) | NAMES `val`: **OVERWRITE** | yes | — | none | same (`srgbClr`) |
| 97 | `crates/litchi-pptx/src/presentation_properties/metadata/protection/codec.rs:138` | `parse_verifier` | chk dflt | FAIL-FAST | NAMES | yes | `set_once` flags | none | `modifyVerifier` |
| 98 | `crates/litchi-pptx/src/presentation_properties/metadata/revision/codec.rs:476` | `root_namespaces` | chk expl | FAIL-FAST | ALL xmlns | yes | HashSet | none | revisionInfo part |
| 99 | `crates/litchi-pptx/src/presentation_properties/metadata/revision/codec.rs:505` | `known_attributes` | chk expl | FAIL-FAST | ALL (allow-list) | yes | HashSet | `known.contains` over a constant slice | same (3 callers) |
| 100 | `crates/litchi-pptx/src/presentation_properties/metadata/revision/codec.rs:544` | `validate_any_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | none | same |
| 101 | `crates/litchi-pptx/src/presentation_properties/metadata/sections/codec.rs:226` | `make_node` | chk expl | FAIL-FAST | ALL | yes | — | **SUPER-LINEAR (same DOM pattern as #83)** | `Sections::from_xml` (plain `Reader`, 8 MiB) |
| 102 | `crates/litchi-pptx/src/presentation_properties/metadata/slide_sync/codec.rs:160` | `from_root` | chk dflt | FAIL-FAST | OTHER (allow-list) | yes | — | none (then 3 fail-fast `unqualified_attribute_value` passes) | `sldSyncPr` part |
| 103 | `crates/litchi-pptx/src/presentation_properties/metadata/structure/codec.rs:732` | `attributes_ns` | chk expl | FAIL-FAST | ALL | yes | HashSet on (ns, local) | ns | presentation structure metadata (3 callers) |
| 104 | `crates/litchi-pptx/src/presentation_properties/metadata/structure/codec.rs:768` | `attributes` (helper) | chk expl | FAIL-FAST | ALL | yes | HashSet\<String\> | none | same (2 callers) |
| 105 | `crates/litchi-pptx/src/presentation_properties/metadata/tracks/tracks_info.rs:523` | `attr` (helper) | chk dflt | FAIL-FAST | NAMES (+ span) | yes | flag | `attribute_span` at most once per call | media track `discover` (14 calls) |
| 106 | `crates/litchi-pptx/src/presentation_properties/readonly_recommended/codec.rs:521` | `parse_value_attr` | chk expl | FAIL-FAST | NAMES `val` (+ span) | yes | flag | `attribute_value_span` at most once | presProps readonlyRecommended ext |
| 107 | `crates/litchi-pptx/src/presentation_properties/readonly_recommended/codec.rs:547` | `extension_uri` | chk expl | FAIL-FAST (returns the first `uri`) | NAMES `uri` | yes | — | none | same |
| 108 | `crates/litchi-pptx/src/shape/classification/codec.rs:585` | `validate_classification_attributes` | chk expl | FAIL-FAST | NAMES `val` | yes | flag | none | p184 shape classification |
| 109 | `crates/litchi-pptx/src/shape/designer/codec.rs:592` | `validate_design_attributes` | chk expl | FAIL-FAST | NAMES `val` | yes | flag | none | `designElem` |
| 110 | `crates/litchi-pptx/src/shape/designer/p202.rs:351` | `parse_editable` | chk expl | FAIL-FAST | NAMES | yes | flag | none | p202 `designPr` |
| 111 | `crates/litchi-pptx/src/shape/designer/p202.rs:382` | `parse_tag` | chk expl | FAIL-FAST | NAMES `name`/`val` | yes | flag (`replace().is_some()`) | none | p202 `designTag` |
| 112 | `crates/litchi-pptx/src/shape/designer/p202.rs:419` | `validate_no_attributes` | chk expl | FAIL-FAST | OTHER | yes | — | none | p202 list elements (3 callers) |
| 113 | `crates/litchi-pptx/src/shape/designer/properties.rs:944` | `extension_uri` | chk expl | FAIL-FAST | NAMES `uri` | yes | flag | none | Designer properties ext |
| 114 | `crates/litchi-pptx/src/shape/designer/properties.rs:968` | `has_extra_attributes` | chk expl | FAIL-FAST (returns at the first extra) | OTHER | yes | — | none | same |
| 115 | `crates/litchi-pptx/src/shape/reader.rs:1423` | `validate_extension_attributes` | chk expl | FAIL-FAST | NAMES `uri` | yes | flag | none (one earlier `unqualified_attribute_value` pass) | shape reader `p:ext` (p232) |
| 116 | `crates/litchi-pptx/src/shape/reader.rs:1479` | `has_forbidden_attribute` | chk expl | FAIL-FAST (returns at the first non-xmlns) | OTHER | yes | — | none | p232 tokens (3 callers) |
| 117 | `crates/litchi-pptx/src/shape/zoom/codec.rs:411` | `apply_namespace_declarations` | chk expl | FAIL-FAST | ALL xmlns | yes | HashMap insert (change log, not a duplicate check) | none | zoom DOM (4 callers) |
| 118 | `crates/litchi-pptx/src/shape/zoom/codec.rs:962` | `dom_frame` | chk expl | FAIL-FAST | ALL | yes | — | none (lexical resolve through a HashMap) | zoom DOM `parse_dom` |
| 119 | `crates/litchi-pptx/src/table/style/codec.rs:180` | `semantic_element` | chk expl | FAIL-FAST (cap 64 non-xmlns) | ALL | yes | — | ns; sort O(n log n) | tableStyles semantic tokens |
| 120 | `crates/litchi-pptx/src/table/style/codec.rs:524` | `attributes` (helper) | chk expl | FAIL-FAST (cap 64) | ALL | yes | — | none | tableStyles (2 callers) |
| 121 | `crates/litchi-pptx/src/tag/codec.rs:241` | `parse_attributes` | chk expl | FAIL-FAST | ALL | yes | HashSet\<String\> | `known.contains` over a constant slice | tags part (3 callers) |
| 122 | `crates/litchi-pptx/src/tag/package/validation.rs:151` | `anchor_relationship_id` | chk expl | FAIL-FAST | NAMES `r:id` (+ span) | yes | flag | ns; `attribute_value_span` at most once | `p:tags` anchors (4 callers) |
| 123 | `crates/litchi-pptx/src/tag/package/validation.rs:197` | `has_non_namespace_attrs` | chk expl | FAIL-FAST (returns at the first) | OTHER | yes | — | none | same (3 callers) |
| 124 | `crates/litchi-pptx/src/tag/shape/codec.rs:1084` | `anchor_id` | chk expl | FAIL-FAST | NAMES `r:id` | yes | flag | ns | shape `p:tags` (2 callers) |
| 125 | `crates/litchi-pptx/src/tag/shape/validation.rs:137` | `has_non_namespace_attrs` | chk expl | FAIL-FAST (returns at the first) | OTHER | yes | — | none | shape `p:tags` |
| 126 | `crates/litchi-pptx/src/transition/reader.rs:477` | `parse_attributes` | chk dflt | FAIL-FAST | NAMES | yes | flags (`reject_duplicate`) | ns (prefixed only) | typed transition `read_with` |
| 127 | `crates/litchi-pptx/src/transition/reader.rs:873` | `bounded_attribute_value` | chk dflt | FAIL-FAST | NAMES | yes | flag | none | preset transition names |
| 128 | `crates/litchi-pptx/src/transition/reader.rs:1035` | `raw_is_portable` | chk dflt | FAIL-FAST | OTHER | yes | — | ns (plus a resolver clone per event, see adjacent findings) | retained transition XML |
| 129 | `crates/litchi-pptx/src/validation.rs:894` | `validate_attributes` | chk expl | FAIL-FAST | NAMES `id`/`r:id` + OTHER | yes | flags | ns | package semantic validator `inspect_xml` (every element) |
| 130 | `crates/litchi-pptx/src/view_properties/codec.rs:196` | `make` | chk expl | FAIL-FAST | ALL | yes | — | **SUPER-LINEAR (same DOM pattern as #83)** | viewProps.xml `ViewProperties::parse` (plain `Reader`, 8 MiB) |

## Helper functions that wrap attribute iteration

Callers inherit each helper's mode. Call counts are grep line counts in production code (`[+ test]` where test code also calls the helper).

**Helpers whose loop is not fail-fast. Every caller is exposed.**

| helper | file:line | mode | prod calls |
|---|---|---|---|
| `SlideBackground::parse_color` | `crates/litchi-pptx/src/backgrounds/codec.rs:92` | SHORT-CIRCUIT ×2 | 4 (`codec.rs:42`, `155`, `248`, `252`) |
| `SlideBackground::parse_gradient_descriptor` | `crates/litchi-pptx/src/backgrounds/codec.rs:185` | TOLERANT ×2, returns `()` | 2 (`:159`, `:164`) |
| `SlideBackground::parse_gradient` | `crates/litchi-pptx/src/backgrounds/codec.rs:128` | TOLERANT | 1 |
| pub `SlideBackground::from_xml` | `crates/litchi-pptx/src/backgrounds/codec.rs:19` | SHORT-CIRCUIT + calls the three above | 0 internal (6 test lines; `crates/litchi/examples/pptx_comprehensive_test.rs`, `crates/litchi-pptx/tests/pptx_effective_backgrounds.rs`) |
| pub `custom_show::List::parse_xml` | `crates/litchi-pptx/src/presentation_properties/metadata/custom_show/codec.rs:18` | TOLERANT ×2 | 0 internal (1 test; `crates/litchi/examples/pptx_isolate_issue.rs`) |
| pub `handout::Master::parse_xml` | `crates/litchi-pptx/src/presentation_properties/metadata/handout/codec.rs:18` | TOLERANT ×2 | 0 internal (1 test; examples, `crates/litchi-pptx/tests/pptx_handout_master.rs`) |
| `inspect_start` | `crates/litchi-pptx/src/animations/package.rs:380` | COUNT (shielded) | 2 |
| `parse_control` | `crates/litchi-pptx/src/presentation/embedded/controls/codec.rs:169` | COUNT | 2 |
| `make_node` (OLE) | `crates/litchi-pptx/src/presentation/embedded/ole/codec.rs:99` | COUNT, then FAIL-FAST | 2 |
| `make_node` (OLE slide) | `crates/litchi-pptx/src/presentation/embedded/ole/slide/codec.rs:158` | COUNT (shielded) | 2 |

**Fail-fast lookup and collection helpers. Callers inherit FAIL-FAST.**

- Single-name lookups:
  - `p14_attribute` (wire.rs:356): 1 call.
  - `relationship_attribute` (animations/package.rs:667): 2 calls, reached through `unqualified_attribute` / `namespaced_attribute` / `required_relationship_attribute`.
  - `element_relationship_id` (master_layout/codec.rs:684): 2.
  - `attribute` (controls/codec.rs:337): 6.
  - `find_attribute` (controls/slide/codec.rs:506): 9.
  - `attr` (tracks_info.rs:515): 14.
  - `bounded_attribute_value` (transition/reader.rs:867): 1.
  - `extension_uri` (readonly_recommended/codec.rs:546): 1.
  - `extension_uri` (designer/properties.rs:938): 1.
  - `element_uri` (guides/codec.rs:352): 2.
  - `black_white_mode` (content_parts/codec.rs:987): 3.
  - `relationship_value_span` (content_parts/codec.rs:856): 2.
  - `anchor_relationship_id` (tag/package/validation.rs:142): 4.
  - `anchor_id` (tag/shape/codec.rs:1075): 2.
  - `requires_value` (ink_actions/codec.rs:672): 1.
  - `required_value` (math/codec.rs:255): 2.
  - `parse_value_attr` (readonly_recommended/codec.rs:513): 1.
- Collectors:
  - `attributes`: chart/extension/codec/xml.rs:514 (3), modern_comments/wire/xml.rs:248 (2), custom_show/wire.rs:741 (2), designer_tags/codec.rs:490 (1), structure/codec.rs:762 (2), table/style/codec.rs:521 (2).
  - `attributes_ns` (structure/codec.rs:725): 3.
  - `collect_unqualified_attributes` (font/codec.rs:394): 2, including one through `reject_unqualified_attributes`.
  - `known_attributes`: modern_comments/codec.rs:316 (1), modern_comments parser.rs:566 (3), changes/codec.rs:888 (5), revision/codec.rs:498 (3).
  - `parse_attributes`: tag/codec.rs:232 (3), transition/reader.rs:461 (2).
  - `capture_relationship_attributes` (ink_actions/codec.rs:736): 2.
  - `semantic_element` (table/style/codec.rs:168): 2.
- Namespace-declaration collectors:
  - `namespace_declarations_from`: modern_comments/codec.rs:366 (2), parser.rs:614 (4).
  - `namespace_declarations`: modern_comments/wire/xml.rs:219 (2), presentation/source.rs:4104 (1).
  - `root_namespaces`: changes/codec.rs:859 (1), revision/codec.rs:469 (1).
  - `apply_namespace_declarations` (zoom/codec.rs:405): 4.
  - `root_fragment_namespace_info` (owner.rs:401): 1.
  - `Namespaces::push` (owner.rs:195): 2.
- Validators and predicates:
  - `check_attribute_count` (animations/codec/validation.rs:294): 4.
  - `validate_any_attributes`: modern_comments/codec.rs:343 (2), parser.rs:593 (6), revision/codec.rs:543 (2).
  - `any_attributes` (changes/codec.rs:932): 2.
  - `no_non_namespace_attributes`: modern_comments/codec.rs:354 (1), parser.rs:604 (3).
  - `no_attributes` (math/codec.rs:288): 2.
  - `validate_attributes`: content_parts/codec.rs:1019 (6), ink_actions/codec.rs:815 (2), validation.rs:883 (2).
  - `has_non_namespace_attrs`: tag/package/validation.rs:196 (3), tag/shape/validation.rs:136 (1).
  - `has_forbidden_attribute` (shape/reader.rs:1475): 3.
  - `has_extra_attributes` (designer/properties.rs:967): 2.
  - `validate_no_attributes` (p202.rs:418): 3.
  - `has_mce_attribute` (order.rs:831): 2.
  - `reject_mce` (opened/copy_plan.rs:1094): 7.
  - `validate_semantic_attributes` (parts/slide.rs:249): 4 [+2 test].
  - `validate_semantic_attribute_names` (parts/slide.rs:276): 4 [+4 test].
  - `Namespaces::preflight` (owner.rs:143): 2.
  - `validate_start_attributes` (owner.rs:1366): 2.
  - `extension_uri_is_supported` (owner.rs:1334): 1.
  - `validate_extension_uri` (placeholder.rs:1978): 2.
  - `validate_extension_attributes` (shape/reader.rs:1417): 1.
  - The remaining sites are single-caller parse functions; the site table has their details.
- DOM node builders (not lookups, but called once per element):
  - `make_node`: comments/codec.rs (2), comments/collaboration/codec.rs (2), media_parts/codec.rs (1), presentation_properties/codec.rs (2), guides/codec.rs (2), sections/codec.rs (2).
  - `make` in view_properties/codec.rs (2).

**Wrappers over `litchi-ooxml-common` helpers. All are FAIL-FAST; each call is one full O(n) pass, so K calls per element cost O(K·n).**

- `litchi_ooxml_common::xml::unqualified_attribute_value` (`crates/litchi-ooxml-common/src/xml.rs:116`, checked, `?` on `Err`, own duplicate flag): about 60 direct call lines in litchi-pptx non-test files.
- `animations::codec::validation::attribute` (`validation.rs:284`) wraps the function above: 61 calls (45 in `xml/parser/semantic.rs`, 16 in `xml/parser/wire.rs`). Elements are already capped at 64 by `check_attribute_count`.
- `litchi_ooxml_common::relationships::attribute_value` (`crates/litchi-ooxml-common/src/relationships/codec.rs:26`, checked, fail-fast, flag, ns) is wrapped by:
  - `crate::namespace::relationship_attribute_value` (`namespace.rs:37`): 8 calls.
  - `crate::parts::relationship_attribute` (`parts/mod.rs:421`): 3 calls (`opened/xml.rs:1558`, `parts/presentation.rs:326`, `:432`).
  - `crate::presentation::embedded::relationship_value` (`presentation/embedded/mod.rs:59`): 5 calls.

## Summary

- **Totals:** 130 sites = 129 production + 1 TEST (`notes/codec.rs:720`, FAIL-FAST). All are checked (89 explicit, 41 default). 0 unchecked, 0 `html_attributes`, 0 NOT-QX.
- **Classes (production):**

| class | count |
|---|---|
| FAIL-FAST | 114 |
| TOLERANT | 7 |
| SHORT-CIRCUIT | 4 |
| COUNT (tolerant `.count()`) | 4 |

  A further 8 FAIL-FAST sites also enforce a count cap on Ok items: validation.rs:296, content_parts:1021, ink_actions:824, owner.rs:146, notes:473 (document-wide), table/style:180 and :524, chart ext:520.

**TOLERANT (all use `.flatten()`, keep iterating after `Err`, Θ(D×R)):**
- `crates/litchi-pptx/src/backgrounds/codec.rs:143`: gradient stop `pos` overwrite, pub.
- `crates/litchi-pptx/src/backgrounds/codec.rs:193`: `lin` `ang` overwrite, returns `()`.
- `crates/litchi-pptx/src/backgrounds/codec.rs:205`: `path` type overwrite, returns `()`.
- `crates/litchi-pptx/src/presentation_properties/metadata/custom_show/codec.rs:31`: custShow `name`/`id` overwrite.
- `crates/litchi-pptx/src/presentation_properties/metadata/custom_show/codec.rs:51`: `sld` ids appended per attribute.
- `crates/litchi-pptx/src/presentation_properties/metadata/handout/codec.rs:28`: `hf` flags overwrite, pub API.
- `crates/litchi-pptx/src/presentation_properties/metadata/handout/codec.rs:49`: handout `srgbClr` `val` overwrite.

**SHORT-CIRCUIT:**
- `crates/litchi-pptx/src/backgrounds/codec.rs:61`: `prst` `find_map`; quadratic when absent or last.
- `crates/litchi-pptx/src/backgrounds/codec.rs:101`: `srgbClr` `val`; quadratic when absent or last.
- `crates/litchi-pptx/src/backgrounds/codec.rs:110`: `schemeClr` `val`; quadratic when absent or last.
- `crates/litchi-pptx/src/presentation/source/svg_lifecycle.rs:1254`: `.next()` only, O(1), harmless.

**COUNT (`.count() > 64`: counts first, counts `Err` items, iterates everything before the cap):**
- `crates/litchi-pptx/src/presentation/embedded/controls/codec.rs:174`: `p:control` elements; **exposed**.
- `crates/litchi-pptx/src/presentation/embedded/ole/codec.rs:105`: every slide element; **exposed**.
- `crates/litchi-pptx/src/animations/package.rs:389`: shielded by the earlier timing parse.
- `crates/litchi-pptx/src/presentation/embedded/ole/slide/codec.rs:167`: shielded by `load_slide` (#57/#58). On its own it would also accept duplicates silently, because the raw `scan_attributes` has no duplicate check.

All 15 non-fail-fast sites are in functions that return `Result`, except `backgrounds/codec.rs:193` and `:205` (`parse_gradient_descriptor` returns `()`; its caller returns `Result`).

**QUADRATIC own-duplicate checks:**
- `crates/litchi-pptx/src/media_parts/codec.rs:577`: Vec `.iter().any` on expanded names, **no count cap** (only the 4 MiB string budget).
- Capped, so bounded:
  - chart/extension/codec/xml.rs:520 (≤64)
  - ink_actions/codec.rs:824 (≤256, two Vecs)
  - svg_lifecycle/owner.rs:1372 (≤256, through `preflight`)
  - presentation/transition.rs:1752 (bounded by the allow-list)

**Other super-linear patterns:**
- `crates/litchi-pptx/src/comments/collaboration/codec.rs:998`: per-attribute raw-tag rescans (`attribute_value_span` + `attribute_full_span`). O(n·L) with no attribute cap and **no duplicates needed**.
- DOM builders `presentation_properties/codec.rs:192`, `guides/codec.rs:607`, `sections/codec.rs:226`, `view_properties/codec.rs:196`:
  - each node clones its parent's `bindings` (time and retained memory O(nodes·B));
  - each xmlns does a linear `find` (k² for k declarations on one tag);
  - each prefixed attribute does a linear `resolve`.
  - Nothing caps B: plain `Reader`, and MCE is skipped when its namespace is absent. For example, a root with ~10⁵ declarations followed by 10⁵ empty children is still under 8 MiB, but it clones about 10¹⁰ String pairs, so allocation fails and the process aborts.
- Outside the attribute loop: `master_layout/package.rs:599` and `:718` scan all earlier entries for every entry, O(E²) with E ≤ 100 000.
- Bounded but per attribute:
  - `ns` resolver scans (O(n·B) over the bindings in scope);
  - one `resolve_prefix` per `Requires` token at ink_actions:678.

**Adjacent findings, out of scope, not verified further:**
- `crates/litchi-ooxml-common/src/mce/codec.rs:337-356` `with_local` (as committed at HEAD `1d1044e3ac`): a Θ(k²) scan for duplicate prefixes (`local[..index].iter().any`) plus a layer walk (`self.get`) for every declaration. Both run before the 4096-binding cap is checked. This hits every pptx parser that runs MCE processing, whenever the document contains the MCE namespace. **Note:** during this survey the worktree gained an uncommitted diff from another agent (`mce/codec.rs`, `mce/model.rs`, new `mce/scope.rs`, and others) that rewrites `with_local`. The litchi-pptx tree itself is unmodified, so every pptx row above reflects HEAD.
- 26 pptx sites call `reader.resolver().clone()` once per event, copying every in-scope namespace byte (URI length is unbounded), so the cost is O(events × namespace bytes). Examples: transition/reader.rs:171, animations wire.rs:401, ole/codec.rs:55, controls/codec.rs:55, content_parts/codec.rs:84/261/439/749.
