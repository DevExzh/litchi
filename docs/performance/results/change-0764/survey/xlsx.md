# litchi-xlsx: quick-xml attribute-iterator call sites (survey for 0764)

Worktree `/home/zhuhe/code/litchi-worktrees/0764-xml-attribute-dos-hardening` at `1d1044e3ac`, quick-xml 0.41.0.
Method: `grep -rnE '\.(html_)?attributes\(\)' crates/litchi-xlsx/src` gives 201 occurrences, none of them `.html_attributes()`.
I read every loop body in full, plus the helpers each loop calls: own duplicate sets, `push_*` helpers, resolvers, span finders.

## Legend

- **checks**: `checked` = `.with_checks(true)`; `checked (default)` = no `with_checks` call; `unchecked` = `.with_checks(false)` (the only `Err`s are syntax errors, never `Duplicated`); `conditional` = a runtime bool; `n/a` = not quick-xml.
- **error mode**:
  - **FAIL-FAST**: the first `Err` ends iteration. It is propagated with `?`, panics with `unwrap` in tests, returns `None` via `.ok()?`, returns `false` via `let Ok .. else`, becomes a `Fallback` error, or ends a `collect::<Result<..>>`.
  - **COUNT**: a loop whose whole job is counting against a cap, or a `.count()` call. `(fail-fast)` = an `Err` is propagated. `(tolerant)` = `.count()` consumes every item, including `Err(Duplicated)`.
  - **SHORT-CIRCUIT**: `.next().is_some()`, which reads only the first item.
  - **TOLERANT**: iteration continues past an `Err`. **No site in this crate is TOLERANT.**
  - **NOT-QX**: not a quick-xml iterator.
- **consumption**: NAMES / ALL / VALIDATE / COUNT / OTHER.
  - `guarded` = a second match of the same name is refused, flagged or ignored.
  - `OVERWRITE` = plain assignment: a later duplicate would replace the earlier value.
  - `first-wins` = the loop returns or breaks at the first match.
- **cap N** = the loop refuses the tag after N attributes. **QUADRATIC** = a linear `Vec` scan per attribute.
- Helper call counts are `grep -c` lines minus the definition line.

## Call sites (201 rows, sorted by file then line)

| file:line | fn | checks | error mode | consumption | can-refuse | own-duplicate-check | other super-linear | reachability |
|---|---|---|---|---|---|---|---|---|
| crates/litchi-xlsx/src/active_x/codec/xml.rs:380 | relationship_ids_in_xml | checked | FAIL-FAST | OTHER: values of r-ns attrs into HashSet | yes | none | none | ActiveX removal: r:id scan of rewritten XML (package/mod.rs:113) |
| crates/litchi-xlsx/src/active_x/codec/xml.rs:568 | make_node | checked | FAIL-FAST | ALL (expanded attrs into DOM node) | yes | Vec `attrs.iter().any` on (ns,local): **QUADRATIC, no count cap** | own check is Θ(n²) over DISTINCT attrs (bounded only by part size) | ActiveX part XML DOM parse |
| crates/litchi-xlsx/src/auto_filter/codec.rs:1226 | unknown_attributes | checked | FAIL-FAST | ALL unknown attrs preserved (`known` is a const list) | yes | none (quick-xml) | none (`push_attribute` O(1), cap 4096) | autoFilter parse (worksheet/table/view); 13 calls |
| crates/litchi-xlsx/src/auto_filter/codec.rs:1813 | optional_attr | checked | FAIL-FAST | NAMES one qname, guarded (2nd -> Err) | yes | per-name guard O(1) | none (13 calls = k passes per tag) | autoFilter parse |
| crates/litchi-xlsx/src/calculation_properties/codec.rs:401 | parse_calc_attributes | checked | FAIL-FAST | NAMES guarded (`seen[slot]`) | yes | `seen` array O(1) | none | workbook.xml calcPr parse (after count-first :591) |
| crates/litchi-xlsx/src/calculation_properties/codec.rs:488 | extension_uri | checked | FAIL-FAST | NAMES `uri` guarded | yes | guard | none | calcPr extLst/ext |
| crates/litchi-xlsx/src/calculation_properties/codec.rs:515 | parse_feature | checked | FAIL-FAST | NAMES `name` guarded | yes | guard | none | calcPr calcFeatures/feature |
| crates/litchi-xlsx/src/calculation_properties/codec.rs:575 | no_attributes | checked | FAIL-FAST | VALIDATE (any non-xmlns -> Err) | yes | n/a | none | calcPr children |
| crates/litchi-xlsx/src/calculation_properties/codec.rs:591 | check_attribute_count | unchecked | COUNT (fail-fast; count-first before every calcPr parse fn; cap `limits.max_attributes()`) | COUNT | yes | none (dups refused by the later checked loops) | none | workbook.xml calcPr elements (6 calls) |
| crates/litchi-xlsx/src/calculation_properties/rewriter.rs:1163 | raw_extension_attributes | checked | FAIL-FAST | NAMES `uri` guarded; others set a flag | yes | guard | none | calcPr rewrite (edit) |
| crates/litchi-xlsx/src/calculation_properties/rewriter.rs:1188 | has_process_content | checked | FAIL-FAST (returns Ok(true) on match) | NAMES presence (mc:ProcessContent) | yes | n/a | none | calcPr rewrite |
| crates/litchi-xlsx/src/calculation_properties/rewriter.rs:1205 | check_attributes | unchecked | COUNT (fail-fast; count-first; cap `limits.max_attributes()`) | COUNT | yes | none | none | calcPr rewrite |
| crates/litchi-xlsx/src/cell_values/validation.rs:482 | scan_attributes | checked | FAIL-FAST | VALIDATE (refuse `*:id`) | yes | n/a | none | value-only cell publication closure check (edit) |
| crates/litchi-xlsx/src/cell_watches.rs:316 | parse_cell_watch_attributes | checked | FAIL-FAST | NAMES `r` guarded | yes | guard | none | worksheet cellWatches parse |
| crates/litchi-xlsx/src/cell_watches.rs:365 | reject_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | cellWatches |
| crates/litchi-xlsx/src/chain/codec.rs:239 | preflight_raw_attributes | checked | FAIL-FAST | VALIDATE sizes | yes | n/a | none | calcChain.xml preflight, every element |
| crates/litchi-xlsx/src/chain/codec.rs:282 | parse_root_attributes | checked | FAIL-FAST | ALL (xmlns decls + extension attrs retained, cap 256 each) | yes | `push_extension_attribute`: Vec `iter().any` on raw qname, QUADRATIC but cap 256 (redundant with quick-xml) | Θ(n²) with n ≤ 256 | calcChain.xml root |
| crates/litchi-xlsx/src/chain/codec.rs:328 | parse_cell | checked | FAIL-FAST | NAMES `set_once` + ALL extension attrs | yes | set_once; `push_extension_attribute` QUADRATIC, cap 256 | Θ(n²) with n ≤ 256, per `<c>` | calcChain.xml every `<c>` |
| crates/litchi-xlsx/src/chain/package.rs:196 | validate_sheet_ids | checked | FAIL-FAST | NAMES `sheetId` guarded | yes | guard | none | workbook.xml `<sheet>` scan in calcChain validation |
| crates/litchi-xlsx/src/chain/tests.rs:554 | unresolved_names | checked (default) | FAIL-FAST (unwrap) | OTHER | no | n/a | none | TEST |
| crates/litchi-xlsx/src/chart_sheet/codec.rs:1163 | make_node | checked | FAIL-FAST | ALL | yes | Vec `iter().any` on (ns,name): **QUADRATIC, no count cap** (only the 4 MiB string budget) | Θ(n²) over distinct attrs | chartsheet XML parse |
| crates/litchi-xlsx/src/chart_sheet/package/codec.rs:418 | make_node | checked | FAIL-FAST | ALL | yes | Vec `iter().any`: **QUADRATIC, no count cap** (4 MiB string budget) | Θ(n²) over distinct attrs | chartsheet package-graph parts |
| crates/litchi-xlsx/src/conditional_formatting/codec.rs:1302 | optional_attr | checked | FAIL-FAST | NAMES by local name, guarded | yes | guard | none (16 calls) | conditionalFormatting parse |
| crates/litchi-xlsx/src/conditional_formatting/package.rs:671 | reject_compatibility_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | CF owner edit |
| crates/litchi-xlsx/src/conditional_formatting/package.rs:774 | validate_owner_element | checked | FAIL-FAST | VALIDATE (const allow-list) | yes | n/a | none | CF owner edit |
| crates/litchi-xlsx/src/conditional_formatting/tests.rs:228 | unresolved_names | checked (default) | FAIL-FAST (unwrap) | OTHER | no | n/a | none | TEST |
| crates/litchi-xlsx/src/connections/codec.rs:452 | preflight_element | unchecked | COUNT (fail-fast; preflight pass; cap `max_attributes`) | COUNT + byte limits | yes | none (the checked `make` at :816 refuses dups) | none | connections.xml preflight, every element |
| crates/litchi-xlsx/src/connections/codec.rs:816 | make | checked | FAIL-FAST (cap `max_attributes`) | ALL (raw qname,value) | yes | none (quick-xml) | none | connections.xml DOM parse |
| crates/litchi-xlsx/src/connections/codec.rs:3666 | fragment_has_default_namespace_declaration | checked | FAIL-FAST (Ok(true) on match) | NAMES presence (`xmlns`) | yes | n/a | none | connections fragment edit |
| crates/litchi-xlsx/src/connections/codec.rs:4889 | has_mce_namespace_declaration | checked | FAIL-FAST | NAMES | yes | n/a | none | TEST |
| crates/litchi-xlsx/src/connections/embedded_data.rs:641 | parse_with_limits | checked | FAIL-FAST (cap `max_xml_attributes`) | NAMES plain assign (id/type/uri) after an expanded-name dup refusal | yes | HashSet<(ns,local)> std SipHash | none | connections.xml parse for Custom Data embedded data |
| crates/litchi-xlsx/src/custom_data/package.rs:1947 | existing_content_type_removals | checked | FAIL-FAST (break at first PartName) | NAMES first-wins | yes | n/a | none | [Content_Types].xml scan (Custom Data removal) |
| crates/litchi-xlsx/src/data_consolidation/codec.rs:276 | parse_consolidation_attributes | checked (default) | FAIL-FAST | NAMES set_once | yes | set_once | none | worksheet dataConsolidate |
| crates/litchi-xlsx/src/data_consolidation/codec.rs:321 | parse_data_refs_attributes | checked (default) | FAIL-FAST | NAMES set_once | yes | set_once | none | dataRefs |
| crates/litchi-xlsx/src/data_consolidation/codec.rs:355 | parse_data_ref_attributes | checked (default) | FAIL-FAST | NAMES set_once | yes | set_once | none | dataRef |
| crates/litchi-xlsx/src/data_type_icons/codec.rs:312 | parse_target_attributes | checked | FAIL-FAST | NAMES `visible` guarded | yes | bool guard | none | data-type icon settings |
| crates/litchi-xlsx/src/data_validation/codec/wire.rs:525 | uid_attr | checked | FAIL-FAST | NAMES xr:uid guarded | yes | guard | none (1 call) | dataValidation parse |
| crates/litchi-xlsx/src/data_validation/codec/wire.rs:555 | optional_attr | checked | FAIL-FAST | NAMES guarded | yes | guard | none (10 calls) | dataValidation parse |
| crates/litchi-xlsx/src/drawing/codec.rs:1037 | check_attribute_lengths | checked (default) | FAIL-FAST | VALIDATE sizes | yes | n/a | none | drawing part parse |
| crates/litchi-xlsx/src/drawing/codec.rs:1144 | required_core_content_part_relationship | checked | FAIL-FAST | NAMES r:id guarded | yes | guard | none | drawing contentPart |
| crates/litchi-xlsx/src/drawing/source.rs:2295 | collect_relationship_attributes | checked | FAIL-FAST | ALL r-namespace attrs into Vec | yes | none | none (own resolver is a HashMap) | drawing source inspection (after preflight cap 256) |
| crates/litchi-xlsx/src/drawing/source.rs:2772 | Namespaces::preflight | checked | COUNT (fail-fast; count-first; cap 256 attrs / 256 decls) | COUNT + VALIDATE | yes | n/a | none | drawing source inspection, every start tag |
| crates/litchi-xlsx/src/drawing/source.rs:2832 | Namespaces::push | checked | FAIL-FAST | ALL xmlns into bindings | yes | none (HashMap `by_prefix`) | none | drawing source |
| crates/litchi-xlsx/src/drawing/source.rs:3027 | namespace_complete_element_fragment | checked | FAIL-FAST | ALL xmlns prefixes into Vec | yes | none | after the loop: `inherited` x `declared.iter().any` | drawing fragment extraction |
| crates/litchi-xlsx/src/drawing/source.rs:3111 | extension_uri | checked | FAIL-FAST | NAMES `uri` guarded | yes | guard | none | drawing a:ext (SVG) |
| crates/litchi-xlsx/src/drawing/source.rs:3166 | relationship_attribute | checked | FAIL-FAST | NAMES guarded | yes | guard | none (2 calls) | drawing source |
| crates/litchi-xlsx/src/drawing/source.rs:3224 | required_core_content_part_relationship | checked | FAIL-FAST | NAMES guarded | yes | guard | none | drawing source |
| crates/litchi-xlsx/src/drawing/source.rs:3281 | relationship_dialect_attribute | checked | FAIL-FAST | NAMES first-wins (`found.is_none()`) | yes | none (first wins) | none (2 calls) | drawing source |
| crates/litchi-xlsx/src/drawing/source.rs:3311 | validate_start_attributes | checked | FAIL-FAST | VALIDATE expanded uniqueness | yes | Vec `seen.iter().any`: QUADRATIC (n ≤ 256 via preflight :2772) | Θ(n²) with n ≤ 256 | drawing source, every start tag |
| crates/litchi-xlsx/src/drawing/source.rs:3496 | bounded_unqualified_attribute | checked | FAIL-FAST (returns first match) | NAMES first-wins | yes | none | none (1 call) | drawing source |
| crates/litchi-xlsx/src/drawing/worksheet_source.rs:619 | relationship_id | checked | FAIL-FAST | NAMES r:id guarded | yes | guard | none | worksheet drawing-reference scan |
| crates/litchi-xlsx/src/drawing/worksheet_source.rs:683 | NamespaceState::push (loop 1) | checked | COUNT (fail-fast; count-first; cap `limits.max_attributes`, default 256) | COUNT + VALIDATE | yes | n/a | none | worksheet source scan, every start tag |
| crates/litchi-xlsx/src/drawing/worksheet_source.rs:741 | NamespaceState::push (loop 2) | checked | FAIL-FAST | ALL xmlns into HashMap | yes | none (HashMap) | none | worksheet source scan |
| crates/litchi-xlsx/src/drawing/worksheet_source.rs:776 | NamespaceState::validate_attributes | checked | **COUNT (tolerant)**: `.count()` consumes every item including Err; runs before loop :778 to size `try_reserve_exact` | COUNT | yes | n/a | **shielded**: runs after push loop :683 (fail-fast, capped) on the same tag, so no Err or duplicate reaches it | worksheet source scan |
| crates/litchi-xlsx/src/drawing/worksheet_source.rs:778 | NamespaceState::validate_attributes | checked | FAIL-FAST | VALIDATE expanded uniqueness | yes | Vec `expanded.iter().any`: QUADRATIC (n ≤ `max_attributes`, default 256) | Θ(n²) with n ≤ 256 | worksheet source scan |
| crates/litchi-xlsx/src/drawing/worksheet_source.rs:880 | validate_declaration | checked | FAIL-FAST (cap 3) | VALIDATE XML decl | yes | state machine | none | worksheet `<?xml?>` declaration |
| crates/litchi-xlsx/src/external_links/codec.rs:2799 | bounded_relationship_attribute_value | checked (default) | FAIL-FAST | NAMES guarded | yes | guard | none (1 call) | externalLink part |
| crates/litchi-xlsx/src/external_links/codec.rs:2856 | bounded_unqualified_attribute_value | checked (default) | FAIL-FAST | NAMES guarded | yes | guard | none (2 calls) | externalLink alternate URLs |
| crates/litchi-xlsx/src/form_control/codec.rs:3008 | parse_root_attributes | checked | FAIL-FAST (cap `max_attributes`, 256) | ALL (fields, unknown attrs, layouts) | yes | none (quick-xml) | `find_attribute_span` rescans the tag from its start for every attribute: Θ(n²·len) with n ≤ 256 | ctrlProp formControlPr parse |
| crates/litchi-xlsx/src/form_control/codec.rs:3164 | parse_item_list_attributes | checked | FAIL-FAST (cap 256) | ALL unknown | yes | none | same `find_attribute_span` cost, Θ(n²) with n ≤ 256 | formControlPr itemLst |
| crates/litchi-xlsx/src/form_control/codec.rs:3226 | parse_item_attributes | checked | FAIL-FAST (cap 256) | NAMES `val` guarded | yes | guard | none | formControlPr item |
| crates/litchi-xlsx/src/form_control/codec.rs:3299 | check_opaque_attribute_budget | checked | COUNT (fail-fast; count-first; cap 256) | COUNT | yes | n/a | none | formControlPr extLst/opaque (3 calls) |
| crates/litchi-xlsx/src/form_control/codec.rs:3321 | extension_list_container_diagnostic | checked | FAIL-FAST | OTHER (flag if any non-xmlns attr) | yes | n/a | none | formControlPr extLst |
| crates/litchi-xlsx/src/form_control/codec.rs:3345 | validate_extension_element | checked | FAIL-FAST | NAMES `uri`; a 2nd one sets `valid=false` | yes | `uri_seen` flag | none | formControlPr ext |
| crates/litchi-xlsx/src/form_control/codec.rs:3420 | namespace_context_for_element (loop 1) | checked | COUNT (fail-fast; count-first: counts xmlns decls to size a reservation) | COUNT | yes | n/a | none | formControlPr itemLst namespace context |
| crates/litchi-xlsx/src/form_control/codec.rs:3429 | namespace_context_for_element (loop 2) | checked | FAIL-FAST | ALL xmlns into Vec | yes | none | none | same |
| crates/litchi-xlsx/src/form_control/model.rs:2402 | validate_opaque_attributes | checked | FAIL-FAST (counts the item, then `?`; cap 256) | VALIDATE | yes | n/a | none | form-control model opaque XML |
| crates/litchi-xlsx/src/form_control/owner.rs:4819 | validate_namespace_declarations | unchecked | FAIL-FAST (syntax Err only) | VALIDATE xmlns values | yes | none | none | form-control owner (worksheet/VML edit) |
| crates/litchi-xlsx/src/form_control/owner.rs:5451 | mce_choice_supported | unchecked | FAIL-FAST (break at first `Requires`) | NAMES first-wins | yes | none | none | owner MCE Choice |
| crates/litchi-xlsx/src/form_control/owner.rs:5582 | reject_mixed_dialect | unchecked | FAIL-FAST (break at first strict r attr) | OTHER presence | yes | none | none | owner dialect scan |
| crates/litchi-xlsx/src/form_control/owner.rs:6663 | parse_control | unchecked | FAIL-FAST | NAMES plain assign: a duplicate shapeId, r:id or name would OVERWRITE (last wins) | yes | none | none | owner worksheet `<control>` |
| crates/litchi-xlsx/src/form_control/owner.rs:6715 | attr_rel_id | unchecked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (4 calls) | owner |
| crates/litchi-xlsx/src/form_control/owner.rs:6941 | preflight_control_relationship_id | unchecked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none | owner |
| crates/litchi-xlsx/src/form_control/owner.rs:7239 | parse_drawing_shape | unchecked | FAIL-FAST | NAMES plain assign: a duplicate id or name would OVERWRITE | yes | none | none | owner DrawingML cNvPr |
| crates/litchi-xlsx/src/form_control/owner.rs:7358 | validate_unique_vml_attributes | unchecked | FAIL-FAST (cap 256) | VALIDATE expanded uniqueness | yes | HashSet<Vec<u8>> std SipHash | none | owner VML shapes |
| crates/litchi-xlsx/src/form_control/owner.rs:9104 | attr_unqualified | unchecked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (6 calls) | owner |
| crates/litchi-xlsx/src/form_control/owner.rs:9128 | attr_qualified | unchecked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (2 calls) | owner |
| crates/litchi-xlsx/src/form_control/scalar_edit.rs:1494 | shape_id_matches | checked | FAIL-FAST | NAMES `id` guarded | yes | guard | none | form-control scalar edit (VML) |
| crates/litchi-xlsx/src/header_footer/codec.rs:232 | parse_settings | checked | FAIL-FAST | NAMES `seen[]` | yes | seen array | none | worksheet headerFooter |
| crates/litchi-xlsx/src/header_footer/codec.rs:262 | validate_child_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | headerFooter children |
| crates/litchi-xlsx/src/hyperlinks/codec.rs:372 | validate_container_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | worksheet hyperlinks |
| crates/litchi-xlsx/src/hyperlinks/codec.rs:409 | push_hyperlink | checked | FAIL-FAST | NAMES plain assign (raw dups refused; two prefixes bound to the r namespace would OVERWRITE) | yes | none | none | worksheet hyperlink |
| crates/litchi-xlsx/src/hyperlinks/codec.rs:653 | validate_exclusive_relationship_references | checked | FAIL-FAST | OTHER (counts r:id references in a HashMap) | yes | n/a | none | worksheet scan in hyperlink edit |
| crates/litchi-xlsx/src/hyperlinks/codec.rs:706 | read_relationship_id | checked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (2 calls) | hyperlinks |
| crates/litchi-xlsx/src/hyperlinks/codec.rs:1242 | root_namespace | checked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none | worksheet root |
| crates/litchi-xlsx/src/ignored_errors/codec.rs:470 | parse_ignored_error | checked | FAIL-FAST | NAMES guarded | yes | `seen_flags` / guard | none | ignoredErrors |
| crates/litchi-xlsx/src/ignored_errors/codec.rs:618 | parse_extension | checked | FAIL-FAST | NAMES guarded | yes | guard | none | ignoredErrors ext |
| crates/litchi-xlsx/src/ignored_errors/codec.rs:649 | reject_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | ignoredErrors |
| crates/litchi-xlsx/src/ignored_errors/tests.rs:175 | unresolved_names | checked (default) | FAIL-FAST (unwrap) | OTHER | no | n/a | none | TEST |
| crates/litchi-xlsx/src/named_sheet_view/codec.rs:1773 | parse_namespace_declarations | checked | FAIL-FAST (cap 256 decls) | ALL xmlns | yes | HashSet<String> std (raw names; redundant) | none | namedSheetView part |
| crates/litchi-xlsx/src/named_sheet_view/codec.rs:1794 | attr | checked | FAIL-FAST | NAMES guarded | yes | guard | none (12 calls) | namedSheetView |
| crates/litchi-xlsx/src/named_sheet_view/codec.rs:2181 | unresolved_names | checked (default) | FAIL-FAST (unwrap) | OTHER | no | n/a | none | TEST |
| crates/litchi-xlsx/src/ole_objects/codec.rs:706 | raw_attribute_value | checked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (2 calls) | worksheet oleObjects |
| crates/litchi-xlsx/src/ole_objects/codec.rs:733 | raw_attributes | checked | FAIL-FAST | ALL (expanded Vec) | yes | none in the loop | **after the loop: `expanded.iter().find` for every source attribute, Θ(n²), no cap** | worksheet raw-source collection for OLE edits, every element |
| crates/litchi-xlsx/src/ole_objects/codec.rs:1216 | make_node | checked | FAIL-FAST | ALL | yes | Vec `iter().any`: **QUADRATIC, no count cap** (4 MiB string budget) | Θ(n²) over distinct attrs | OLE markup DOM parse |
| crates/litchi-xlsx/src/outline_properties.rs:320 | parse_attributes | checked | FAIL-FAST | NAMES set_once | yes | set_once | none | worksheet outlinePr |
| crates/litchi-xlsx/src/outline_properties.rs:368 | validate_attribute_syntax | checked | FAIL-FAST | VALIDATE | yes | n/a | none | outlinePr |
| crates/litchi-xlsx/src/page_breaks/codec.rs:529 | reject_unknown_attribute_prefixes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | rowBreaks/colBreaks |
| crates/litchi-xlsx/src/page_breaks/codec.rs:550 | collection_metadata | checked | FAIL-FAST | OTHER (lossy flag) | yes | n/a | none | page breaks |
| crates/litchi-xlsx/src/page_breaks/codec.rs:616 | begin_collection | checked | FAIL-FAST | NAMES guarded | yes | guard | none | page breaks |
| crates/litchi-xlsx/src/page_breaks/codec.rs:661 | parse_break | checked | FAIL-FAST | NAMES set_once | yes | set_once | none | `<brk>` |
| crates/litchi-xlsx/src/page_margins.rs:358 | parse_margins | checked | FAIL-FAST | NAMES guarded | yes | `values[]` guard | none | pageMargins |
| crates/litchi-xlsx/src/page_setup/codec.rs:397 | parse_setup | checked | FAIL-FAST | NAMES `seen[]` | yes | seen array | none | pageSetup |
| crates/litchi-xlsx/src/phonetic_properties.rs:307 | parse_attributes | checked | FAIL-FAST | NAMES set_once | yes | set_once | none | phoneticPr |
| crates/litchi-xlsx/src/pivot/server_formats/cached_unique_names.rs:68 | scan_mce_branch | checked | FAIL-FAST | NAMES `Requires` guarded | yes | guard | each attr calls `scope_contains`, which walks the ignorable scopes (n x M ignorable namespaces) | pivot server-formats MCE Choice |
| crates/litchi-xlsx/src/pivot/server_formats/cached_unique_names.rs:164 | scan_mce_branch | checked | FAIL-FAST | VALIDATE | yes | n/a | same `scope_contains` cost, n x M | MCE Fallback |
| crates/litchi-xlsx/src/pivot/server_formats/cached_unique_names.rs:206 | validate_mce_alternate_content_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | same `scope_contains` cost, n x M | MCE AlternateContent |
| crates/litchi-xlsx/src/pivot/server_formats/mod.rs:5979 | begin_namespace_scope | checked | FAIL-FAST (cap `namespace_declarations`) | ALL xmlns into resolver | yes | none (HashMap resolver/interner) | none | pivot part scan, every start tag |
| crates/litchi-xlsx/src/pivot/server_formats/mod.rs:6050 | parse_ignorable_namespaces | checked | FAIL-FAST | NAMES mc:Ignorable (Ignorables under different MCE-bound prefixes add up) | yes | Vec `namespaces.iter().any` per token (≤ 256) | per-token dedup O(256) + resolve | pivot part |
| crates/litchi-xlsx/src/pivot/server_formats/mod.rs:6212 | parse_attrs | checked | FAIL-FAST (cap 256 non-xmlns) | ALL | yes | Vec `attrs.iter().any`: QUADRATIC (n ≤ 256) | Θ(n²) with n ≤ 256 | pivot part, every element |
| crates/litchi-xlsx/src/print_options.rs:355 | parse_options | checked | FAIL-FAST | NAMES `seen[]` | yes | seen array | none | printOptions |
| crates/litchi-xlsx/src/query_table/codec.rs:201 | make_node | checked | FAIL-FAST | ALL | yes | HashSet<String> std (raw qname; redundant) | none | queryTable part DOM |
| crates/litchi-xlsx/src/raw/catalog_edit/codec.rs:712 | tag | checked (default) | FAIL-FAST | ALL (copied for re-serialization) | yes | none | none | workbook.xml catalog edits |
| crates/litchi-xlsx/src/raw/compact.rs:204 | compact_events | checked | FAIL-FAST | NAMES `xml:space` plain assign (dups refused) | yes | none | none | compact publication of changed parts, every start tag |
| crates/litchi-xlsx/src/raw/compact.rs:344 | write_start | checked | FAIL-FAST | ALL (copied into a normalized tag) | yes | none | none | compact publication write |
| crates/litchi-xlsx/src/raw/compact.rs:404 | changed_owned | checked | FAIL-FAST | NAMES | yes | none | none | TEST |
| crates/litchi-xlsx/src/raw/properties_edit.rs:653 | tag | checked (default) | FAIL-FAST | ALL copy | yes | none | none | docProps/app.xml edit (sheet arrange/remove) |
| crates/litchi-xlsx/src/raw/reference_edit.rs:363 | attribute_replacement | checked (default) | FAIL-FAST | ALL scanned (sheet-name and formula rewrites) | yes | none | each attr: `direct_name` over renames + formula rename (n x R); plus one `relationship_attribute_value` pass before the loop | sheet rename across parts |
| crates/litchi-xlsx/src/raw/reference_edit.rs:617 | tag | checked (default) | FAIL-FAST | ALL copy | yes | none | none | sheet rename |
| crates/litchi-xlsx/src/raw/reference_scan.rs:218 | scan_attributes | checked (default) | FAIL-FAST (returns on first hit) | ALL scanned | yes | none | each attr: scan all sheets / `depends_on_sheet` (n x S); plus one `relationship_attribute_value` pass before the loop | sheet-dependency scan |
| crates/litchi-xlsx/src/raw/sheet_view_edit.rs:631 | tag | checked (default) | FAIL-FAST | ALL copy | yes | none | none | sheet-view selection sync edit |
| crates/litchi-xlsx/src/raw/web.rs:549 | attribute | checked (default) | FAIL-FAST | NAMES guarded | yes | guard | none (5 calls) | worksheet web-extension bindings |
| crates/litchi-xlsx/src/raw/web.rs:567 | reject_attributes | checked (default) | FAIL-FAST | VALIDATE | yes | n/a | none | web-extension bindings |
| crates/litchi-xlsx/src/raw/web.rs:577 | reject_other_attributes | checked (default) | FAIL-FAST | VALIDATE | yes | n/a | none | x15:webExtension |
| crates/litchi-xlsx/src/raw/worksheet/codec.rs:236 | scan_cell_attributes | checked | FAIL-FAST | NAMES plain assign (dups refused) | yes | none | none | worksheet parse, every `<c>` (hot path) |
| crates/litchi-xlsx/src/raw/worksheet/codec.rs:538 | open_lane_cell | n/a | NOT-QX (`lane::Tag::attributes`, our own iterator; admitted tags have ≤ 32 attrs and duplicate names are declined) | NAMES | yes | n/a | none | lane fast path |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs:617 | raw_attribute | unchecked | FAIL-FAST (`.ok()?` returns None) | NAMES one name; a 2nd occurrence returns None | no (Option; None = ineligible) | `found.is_some()` guard O(1) | none (4 calls, 2 via `raw_u32`) | worksheet edit snapshot facts |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs:637 | lane_cell | n/a | NOT-QX (`lane::Tag::attributes`) | NAMES + count | yes | n/a | none | lane fast path |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs:674 | lane_cell | n/a | NOT-QX (`lane::Tag::attributes`) | ALL | yes | n/a | none | lane fast path |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs:1718 | shared_formula_attributes_supported | checked (default) | FAIL-FAST | NAMES bool flags (dup returns Ok(false)) | yes | bools | none | worksheet edit snapshot `<f>` |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs:78 | tag | checked (default) | FAIL-FAST | ALL copy | yes | none | none | worksheet edit wire |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs:112 | cell_tag | checked (default) | FAIL-FAST | ALL copy | yes | none | none | worksheet edit `<c>` |
| crates/litchi-xlsx/src/raw/worksheet/lane.rs:976 | iterates_raw_attributes | n/a | NOT-QX | OTHER | no | n/a | none | TEST |
| crates/litchi-xlsx/src/raw/worksheet/lane.rs:990 | iterates_raw_attributes | n/a | NOT-QX | COUNT | no | n/a | none | TEST |
| crates/litchi-xlsx/src/raw/worksheet/mod.rs:207 | observe_start | checked | FAIL-FAST (`let Ok .. else return false`) | VALIDATE | no (bool; false = fall back to the full path) | n/a | each MCE Ignorable attr runs `valid_ignorable_directive`: O(T²) token dedup + bindings scan (T effectively ≤ in-scope bindings; cap 4096) | worksheet MCE-preprocessing eligibility |
| crates/litchi-xlsx/src/raw/worksheet/selected.rs:824 | plain_attributes | conditional (checked only if the raw tag contains `xmlns`) | FAIL-FAST (becomes Fallback; cap `max_attributes_per_event`) | NAMES: one allowed name (`r`/`ref`); a 2nd one falls back | yes | `seen_reference` flag | none | selected-range streaming read |
| crates/litchi-xlsx/src/raw/worksheet/validation.rs:107 | validate_defaults_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | worksheet sheetFormatPr |
| crates/litchi-xlsx/src/raw/worksheet/x14ac.rs:709 | descent | checked | FAIL-FAST | NAMES guarded (`replace`) | yes | guard | none (2 calls) | worksheet rows x14ac:dyDescent |
| crates/litchi-xlsx/src/raw/worksheet/x14ac.rs:740 | attribute_name | checked | FAIL-FAST | NAMES guarded | yes | guard | none (4 calls) | worksheet edit snapshot rows |
| crates/litchi-xlsx/src/revisions/codec.rs:469 | make_node | checked | FAIL-FAST | ALL | yes | HashSet<(ns,local)> std | none | revision-log parts |
| crates/litchi-xlsx/src/rich_values/codec/xml.rs:249 | make_node | checked | FAIL-FAST | ALL | yes | Vec `iter().any`: **QUADRATIC, no count cap** (1 MiB string budget) | Θ(n²) over distinct attrs | rich-value parts DOM |
| crates/litchi-xlsx/src/rich_values/refresh_intervals/codec.rs:867 | type_name | checked | FAIL-FAST | NAMES guarded | yes | guard | none | rich-value refresh intervals |
| crates/litchi-xlsx/src/rich_values/refresh_intervals/codec.rs:905 | parse_interval_attributes | checked | FAIL-FAST | NAMES guarded (`is_none`) | yes | guards | none | refreshInterval |
| crates/litchi-xlsx/src/rich_values/refresh_intervals/codec.rs:945 | no_non_namespace_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none (4 calls) | refresh intervals |
| crates/litchi-xlsx/src/row_visibility/rewrite.rs:280 | parse_row | checked | FAIL-FAST | NAMES `r` guarded | yes | guard | none | row-visibility rewrite (edit) |
| crates/litchi-xlsx/src/scenarios/codec.rs:466 | parse_scenarios_attributes | checked (default) | FAIL-FAST | NAMES set_once + ALL unknown preserved | yes | set_once | none (`push_attribute` O(1), cap 65,536) | worksheet scenarios |
| crates/litchi-xlsx/src/scenarios/codec.rs:504 | parse_scenario_attributes | checked (default) | FAIL-FAST | NAMES set_once + ALL unknown preserved | yes | set_once | none | `<scenario>` |
| crates/litchi-xlsx/src/scenarios/codec.rs:558 | parse_input_cell_attributes | checked (default) | FAIL-FAST | NAMES set_once + ALL unknown preserved | yes | set_once | none | `<inputCells>` |
| crates/litchi-xlsx/src/scenarios/codec.rs:698 | add_namespace_declarations | checked (default) | FAIL-FAST | ALL xmlns names into Vec | yes | none | after the loop: in-scope bindings x `declared.iter().any` | unknown scenario element capture |
| crates/litchi-xlsx/src/sheet/properties.rs:545 | parse_sheet_attributes | checked | FAIL-FAST | NAMES set_once | yes | set_once | none | sheetPr |
| crates/litchi-xlsx/src/sheet/properties.rs:627 | parse_tab_color | checked | FAIL-FAST | NAMES set_once | yes | set_once | none | tabColor |
| crates/litchi-xlsx/src/sheet/properties.rs:684 | parse_page_setup_properties | checked | FAIL-FAST | NAMES set_once | yes | set_once | none | pageSetUpPr |
| crates/litchi-xlsx/src/sheet/sparklines.rs:1041 | extra_group_attributes | checked (default) | FAIL-FAST | ALL unknown extras | yes | none | none (`KNOWN` is a const list) | x14 sparklineGroup |
| crates/litchi-xlsx/src/sheet_calculation_properties.rs:265 | parse_sheet_calc_pr_attributes | checked | FAIL-FAST | NAMES guarded | yes | guard | none | sheetCalcPr |
| crates/litchi-xlsx/src/sheet_protection/codec.rs:427 | parse_sheet_protection | checked (default) | FAIL-FAST | NAMES (plain assign behind `seen`) | yes | HashSet<Vec<u8>> std (local names) | none | sheetProtection |
| crates/litchi-xlsx/src/sheet_protection/codec.rs:490 | parse_pending_range | checked (default) | FAIL-FAST | NAMES set_once | yes | HashSet<Vec<u8>> std | none | protectedRange |
| crates/litchi-xlsx/src/sheet_protection/codec.rs:1290 | attribute | checked (default) | FAIL-FAST | NAMES set_once | yes | set_once | none (2 calls) | sheet protection |
| crates/litchi-xlsx/src/sheet_view/codec.rs:901 | attr | checked | FAIL-FAST | NAMES guarded | yes | guard | none (16 calls) | sheetViews |
| crates/litchi-xlsx/src/sheet_view/tests.rs:158 | unresolved_names | checked (default) | FAIL-FAST (unwrap) | OTHER | no | n/a | none | TEST |
| crates/litchi-xlsx/src/slicer_cache.rs:541 | parse_root_attributes | checked | FAIL-FAST | NAMES set_once + ALL retained (cap 128) | yes | set_once | none | slicerCache part root |
| crates/litchi-xlsx/src/slicer_cache.rs:593 | parse_pivot_table | checked | FAIL-FAST | NAMES set_once + ALL retained (cap 128) | yes | set_once | none | slicerCache pivotTable |
| crates/litchi-xlsx/src/slicer_cache.rs:636 | reject_attributes | checked | FAIL-FAST | VALIDATE (const allow-list) | yes | n/a | none | slicerCache |
| crates/litchi-xlsx/src/slicer_cache.rs:1018 | validate_element_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | slicerCache elements |
| crates/litchi-xlsx/src/slicer_cache/crud.rs:1138 | attribute_value | checked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (2 calls) | slicer CRUD edit |
| crates/litchi-xlsx/src/slicer_cache/package.rs:531 | attribute | checked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (4 calls) | slicer package |
| crates/litchi-xlsx/src/slicer_cache/views.rs:599 | parse_known_attributes | checked | FAIL-FAST | NAMES known into Vec + extensions (cap 128) | yes | none | none (`known` is a const list) | slicer parts |
| crates/litchi-xlsx/src/smart_tags/codec.rs:466 | parse_cell | checked | FAIL-FAST | NAMES guarded | yes | guard | none | cellSmartTags |
| crates/litchi-xlsx/src/smart_tags/codec.rs:497 | parse_tag | checked | FAIL-FAST | NAMES guarded | yes | guards | none | cellSmartTag |
| crates/litchi-xlsx/src/smart_tags/codec.rs:544 | parse_property | checked | FAIL-FAST | NAMES guarded | yes | guards | none | cellSmartTagPr |
| crates/litchi-xlsx/src/smart_tags/codec.rs:578 | reject_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | smartTags |
| crates/litchi-xlsx/src/style/stylesheet/parser.rs:953 | parse_alignment | checked (default) | FAIL-FAST | NAMES bitmask (`mark_property`) | yes | bitmask | none | styles.xml alignment |
| crates/litchi-xlsx/src/survey.rs:2005 | validate_element_names | checked | FAIL-FAST | VALIDATE expanded uniqueness | yes | HashSet std (expanded names) | none | survey part, every element |
| crates/litchi-xlsx/src/survey.rs:2042 | attrs | checked | FAIL-FAST (`collect::<Result<Vec<_>,_>>`) | ALL | yes | none | none (4 calls) | survey |
| crates/litchi-xlsx/src/survey.rs:2057 | raw_attributes | checked | FAIL-FAST (collect into Result) | ALL | yes | none | none (2 calls + 4 through `unknown_attributes`) | survey |
| crates/litchi-xlsx/src/tab_state/source.rs:540 | observe_element | checked (default) | FAIL-FAST | NAMES `lockStructure` (OR-accumulated; dups refused) | yes | none | none | workbook.xml tab-state audit |
| crates/litchi-xlsx/src/timelines/codec.rs:546 | make_node | checked | FAIL-FAST | ALL | yes | Vec `iter().any`: **QUADRATIC, no count cap** (1 MiB string budget) | Θ(n²) over distinct attrs | timeline parts DOM |
| crates/litchi-xlsx/src/volatile_dependencies/codec.rs:651 | optional_attr | checked | FAIL-FAST | NAMES guarded | yes | guard | none (2 calls) | volatileDependencies.xml |
| crates/litchi-xlsx/src/volatile_dependencies/codec.rs:668 | only_attrs | checked | FAIL-FAST | VALIDATE | yes | n/a | none (5 calls) | volatileDependencies.xml |
| crates/litchi-xlsx/src/volatile_dependencies/codec.rs:686 | namespace_attributes | checked | FAIL-FAST | ALL xmlns | yes | none | none (2 calls) | volatileDependencies.xml |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:1433 | rewrite_load_version | checked | FAIL-FAST (returns at first match) | NAMES first-wins | yes | none | none | Data Model load-version rewrite |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:1541 | extension_uri | checked | FAIL-FAST (returns first; local-name match) | NAMES first-wins | yes | none | none (4 calls) | workbook.xml ext |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:1590 | validate_opaque_element | checked | **COUNT (tolerant)**: `.count()` consumes every item including `Err(Duplicated)`; runs first, to size a HashSet `try_reserve` | COUNT | yes | n/a | **UNSHIELDED**: first checked pass over the tag (NsReader's push is unchecked), so Θ(D×R); tag ≤ 4 MiB (`MAX_EXTENSION_BYTES`), document ≤ 16 MiB | workbook.xml elements inside the x15 dataModel `extLst` opaque subtree; Data Model load / rewrite (`load_data_model`, Snapshot) |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:1595 | validate_opaque_element | checked | FAIL-FAST | VALIDATE expanded uniqueness | yes | HashSet<(ns,local)> std | none | same |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:1872 | make_node | checked | FAIL-FAST | ALL | yes | Vec `iter().any`: **QUADRATIC, no count cap** (16 MiB string budget) | Θ(n²) over distinct attrs | workbook.xml DOM parse (every non-opaque element) on Data Model load/removal |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:1944 | opaque_capture | checked | FAIL-FAST | ALL xmlns prefixes into HashSet | yes | HashSet<&[u8]> std | none | extLst capture |
| crates/litchi-xlsx/src/workbook/data_model/codec.rs:2364 | xml_attribute_range | checked | FAIL-FAST (returns first) | NAMES first-wins | yes | none | none (1 call) | Data Model edit |
| crates/litchi-xlsx/src/workbook/data_model/removal.rs:310 | prepare | checked | FAIL-FAST | NAMES `id` (dups refused) | yes | none | none (`deleting` is a HashSet) | Data Model removal |
| crates/litchi-xlsx/src/workbook/data_model/removal.rs:538 | check_formulas | checked | FAIL-FAST | VALIDATE | yes | n/a | each attr recomputes the loop-invariant `resolve_element(element.name())`, O(bindings) | Data Model removal formula scan |
| crates/litchi-xlsx/src/workbook/edit/drawing_transfer.rs:702 | parse_anchor | checked | FAIL-FAST | ALL r-ns values into HashSet | yes | none | none | cross-workbook drawing transfer |
| crates/litchi-xlsx/src/workbook/edit/drawing_transfer.rs:986 | inspect_worksheet_child | checked | FAIL-FAST | NAMES plain assign `drawing_reference` (two r-ns prefixes would OVERWRITE) | yes | none | none | drawing-transfer worksheet scan |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:4987 | manifest_svg_inventory | checked | FAIL-FAST | NAMES plain assign by local name (prefixed variants would OVERWRITE) | yes | none | none | [Content_Types].xml (SVG lifecycle edit) |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:5200 | ext_list_contains_only_owner | checked (default) | SHORT-CIRCUIT (`.next().is_some()` reads the first item only; an Err counts as present) | OTHER presence | yes | n/a | none, O(1); never reaches a duplicate | SVG ext-list check (edit) |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:5217 | ext_list_contains_only_owner | checked (default) | SHORT-CIRCUIT (`.next().is_some()`) | OTHER presence | yes | n/a | none, O(1) | SVG ext-list check (edit) |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:5278 | ext_list_contains_only_owner_ranges | checked (default) | SHORT-CIRCUIT (`.next().is_some()`) | OTHER presence | yes | n/a | none, O(1) | SVG ext-list check (edit) |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:5292 | ext_list_contains_only_owner_ranges | checked (default) | SHORT-CIRCUIT (`.next().is_some()`) | OTHER presence | yes | n/a | none, O(1) | SVG ext-list check (edit) |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:5523 | manifest_has_override | checked | FAIL-FAST | NAMES plain assign by local name | yes | none | none | [Content_Types].xml |
| crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs:5582 | manifest_has_default_svg | checked | FAIL-FAST | NAMES plain assign by local name | yes | none | none | [Content_Types].xml |
| crates/litchi-xlsx/src/workbook/edit/transfer.rs:567 | validate_strict_attributes | checked | FAIL-FAST | VALIDATE | yes | n/a | none | strict worksheet transfer |
| crates/litchi-xlsx/src/workbook/source_merge.rs:1285 | attribute_value | unchecked | FAIL-FAST (syntax Err only; returns first) | NAMES first-wins by local name (prefix-insensitive) | yes | none | none (5 calls) | worksheet merged-range edit |
| crates/litchi-xlsx/src/workbook_metadata/codec.rs:608 | node | checked | FAIL-FAST | ALL | yes | Vec `values.iter().any`: **QUADRATIC, no count cap** | Θ(n²) over distinct attrs | metadata.xml DOM |
| crates/litchi-xlsx/src/workbook_metadata/protection.rs:337 | parse_protection_element | checked (default) | FAIL-FAST | NAMES plain assign behind `seen` | yes | HashSet<Vec<u8>> std | none | workbook.xml workbookProtection |

## Helper functions that wrap attribute iteration

Every caller inherits the helper's mode. **All helpers are FAIL-FAST; none is tolerant.**

| helper (definition) | checks / mode | calls |
|---|---|---|
| auto_filter/codec.rs:1811 `optional_attr` | checked, FAIL-FAST, guarded | 13 |
| auto_filter/codec.rs:1220 `unknown_attributes` | checked, FAIL-FAST, ALL unknown | 13 |
| conditional_formatting/codec.rs:1300 `optional_attr` | checked, FAIL-FAST, guarded (local name) | 16 |
| data_validation/codec/wire.rs:549 `optional_attr` (pub(crate)) | checked, FAIL-FAST, guarded | 10 (module) |
| data_validation/codec/wire.rs:519 `uid_attr` | checked, FAIL-FAST, guarded | 1 |
| drawing/source.rs:3159 `relationship_attribute` | checked, FAIL-FAST, guarded | 2 |
| drawing/source.rs:3275 `relationship_dialect_attribute` | checked, FAIL-FAST, first-wins | 2 |
| drawing/source.rs:3489 `bounded_unqualified_attribute` | checked, FAIL-FAST, first-wins | 1 |
| external_links/codec.rs:2791 `bounded_relationship_attribute_value` | checked (default), FAIL-FAST, guarded | 1 |
| external_links/codec.rs:2849 `bounded_unqualified_attribute_value` | checked (default), FAIL-FAST, guarded | 2 |
| form_control/owner.rs:6709 `attr_rel_id` | unchecked, FAIL-FAST, first-wins | 4 |
| form_control/owner.rs:9096 `attr_unqualified` | unchecked, FAIL-FAST, first-wins | 6 |
| form_control/owner.rs:9119 `attr_qualified` | unchecked, FAIL-FAST, first-wins | 2 |
| hyperlinks/codec.rs:701 `read_relationship_id` | checked, FAIL-FAST, first-wins | 2 |
| named_sheet_view/codec.rs:1792 `attr` | checked, FAIL-FAST, guarded | 12 |
| ole_objects/codec.rs:700 `raw_attribute_value` | checked, FAIL-FAST, first-wins | 2 |
| raw/web.rs:543 `attribute` | checked (default), FAIL-FAST, guarded | 5 |
| raw/worksheet/edit/codec/snapshot/facts.rs:616 `raw_attribute` (wrapped by `raw_u32` :641, 2 calls) | unchecked, FAIL-FAST (None), guarded (None) | 4 |
| raw/worksheet/x14ac.rs:703 `descent` | checked, FAIL-FAST, guarded | 2 |
| raw/worksheet/x14ac.rs:735 `attribute_name` | checked, FAIL-FAST, guarded | 4 |
| sheet_protection/codec.rs:1283 `attribute` | checked (default), FAIL-FAST, set_once | 2 |
| sheet_view/codec.rs:899 `attr` | checked, FAIL-FAST, guarded | 16 |
| slicer_cache/crud.rs:1133 `attribute_value` | checked, FAIL-FAST, first-wins | 2 |
| slicer_cache/package.rs:530 `attribute` | checked, FAIL-FAST, first-wins | 4 |
| volatile_dependencies/codec.rs:649 `optional_attr` / :667 `only_attrs` / :684 `namespace_attributes` | checked, FAIL-FAST | 2 / 5 / 2 |
| workbook/data_model/codec.rs:1537 `extension_uri` | checked, FAIL-FAST, first-wins | 4 |
| workbook/data_model/codec.rs:2357 `xml_attribute_range` | checked, FAIL-FAST, first-wins | 1 |
| workbook/source_merge.rs:1280 `attribute_value` | unchecked, FAIL-FAST, first-wins by local name | 5 |
| survey.rs:2040 `attrs` / :2055 `raw_attributes` (wrapped by `unknown_attributes` :2076, 4 calls) | checked, FAIL-FAST (collect Result), ALL | 4 / 2 |
| calculation_properties/codec.rs:589 `check_attribute_count` | unchecked, COUNT (fail-fast) | 6 |
| form_control/codec.rs:3298 `check_opaque_attribute_budget` | checked, COUNT (fail-fast) | 3 |
| rich_values/refresh_intervals/codec.rs:940 `no_non_namespace_attributes` | checked, FAIL-FAST | 4 |
| `reject_attributes` in cell_watches.rs:364 / ignored_errors/codec.rs:648 / smart_tags/codec.rs:577 / raw/web.rs:566 / slicer_cache.rs:630 | checked, FAIL-FAST | 1 / 2 / 1 / 2 / 2 |
| raw/namespace.rs:33 `relationship_attribute_value` (pub), which calls out-of-crate `litchi_ooxml_common::relationships::attribute_value` | checked (default), FAIL-FAST, guarded | 13 (8 files) |
| out of crate, not in the table: `litchi_ooxml_common::xml::unqualified_attribute_value` | checked (default), FAIL-FAST, guarded | 97 call lines in litchi-xlsx |

## Summary

**Counts**

- 201 occurrences: 189 production quick-xml sites, 7 quick-xml test sites, 5 NOT-QX (`lane::Tag::attributes`, 2 of them tests).
- Checks on the 189 production sites: 173 checked (explicit or default), 15 unchecked, 1 conditional.
- Classes on the 189 production sites:

| class | sites |
|---|---|
| FAIL-FAST | 176 |
| COUNT (fail-fast) | 7 |
| COUNT (tolerant) | 2 |
| SHORT-CIRCUIT | 4 |
| TOLERANT | 0 |

- All 7 test sites are FAIL-FAST: 5 use `unwrap`, 2 use `?`.

**Production TOLERANT sites:** none. No `.flatten()`, `.filter_map(..ok)`, `if let Ok` or `Err(_) => continue` exists anywhere in the crate.

**Production COUNT sites**

| site | behaviour |
|---|---|
| workbook/data_model/codec.rs:1590 | `.count()`, **tolerant and UNSHIELDED**: first checked pass over an opaque x15 extLst tag, so Θ(D×R) on a tag of up to 4 MiB |
| drawing/worksheet_source.rs:776 | `.count()`, tolerant but shielded: runs after the fail-fast, capped loop :683 on the same tag |
| calculation_properties/codec.rs:591 | fail-fast, unchecked, count-first cap |
| calculation_properties/rewriter.rs:1205 | fail-fast, unchecked, count-first cap |
| connections/codec.rs:452 | fail-fast, unchecked preflight cap |
| drawing/source.rs:2772 | fail-fast preflight, cap 256 |
| drawing/worksheet_source.rs:683 | fail-fast, count-first, cap 256 |
| form_control/codec.rs:3299 | fail-fast, count-first, cap 256 |
| form_control/codec.rs:3420 | fail-fast, counts xmlns declarations to size a reservation |

**Production SHORT-CIRCUIT sites:** workbook/edit/svg_lifecycle.rs:5200, 5217, 5278 and 5292. Each is `.next().is_some()`, which reads only the first attribute, so it is O(1) and never quadratic.

**QUADRATIC own duplicate checks**

These are our code, not quick-xml. They cost Θ(n²) over *distinct* attributes, so no duplicates are needed, and quick-xml's hash path is O(n) for these tags.

- **No per-tag attribute cap** (bounded only by part size or string budget). Each is a `Vec::iter().any` check on the expanded name in a DOM builder:
  - active_x/codec/xml.rs:568
  - chart_sheet/codec.rs:1163
  - chart_sheet/package/codec.rs:418
  - ole_objects/codec.rs:1216
  - rich_values/codec/xml.rs:249
  - timelines/codec.rs:546
  - workbook/data_model/codec.rs:1872
  - workbook_metadata/codec.rs:608
  - also ole_objects/codec.rs:733, where the pass after the loop runs `expanded.iter().find` for every source attribute, on every worksheet element.
- **Capped at 256, so bounded:**
  - chain/codec.rs:282 and :328 (`push_extension_attribute`, raw-name check that is redundant with quick-xml)
  - drawing/source.rs:3311
  - drawing/worksheet_source.rs:778
  - pivot/server_formats/mod.rs:6212

**Other super-linear patterns**

- form_control/codec.rs:3008 and :3164: `find_attribute_span` rescans the tag for every attribute, Θ(n²·len) with n ≤ 256.
- After-loop cross-scans of in-scope bindings against declarations: drawing/source.rs:3027 and scenarios/codec.rs:698.
- Per-attribute `scope_contains` walk: pivot/server_formats/cached_unique_names.rs:68, :164 and :206.
- Per-token `iter().any` dedup in pivot/server_formats/mod.rs:6050.
- raw/worksheet/mod.rs:207: `valid_ignorable_directive`, O(T²) per MCE Ignorable attribute.
- Attribute × sheet or rename scans: raw/reference_scan.rs:218 and raw/reference_edit.rs:363.
- Cross-cutting: quick-xml `NamespaceResolver::resolve_attribute` is a reverse linear search of the in-scope bindings. Every per-attribute resolve therefore costs O(B), where B is capped only at 256 declarations per element (quick-xml default) and accumulates with depth.

**Correctness notes (not DoS)**

- Unchecked plain-assign loops silently keep the *last* duplicate: form_control/owner.rs:6663 (`parse_control`) and :7239 (`parse_drawing_shape`).
- Several checked loops match by local name or by resolved namespace, so two different prefixes can overwrite each other: hyperlinks/codec.rs:409, drawing_transfer.rs:986, svg_lifecycle.rs:4987, 5523 and 5582.
- The first-wins helpers never see a later duplicate of the name they return.
