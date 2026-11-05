# 0362ul94cb — Custom XML bounded projection follow-up

This record documents the 94cb follow-up returned from the 94ca Custom XML
lossless-rewrite audit.  The implementation is committed at
`e1e7817bd006f3261da5d1a3f978805d63fc5b9e` with tree
`c60d32e5947b3c54ae6691b2703176beecbd393b`; that tip remains immutable for
independent review.  This record makes no acceptance claim.

The final implementation keeps source-backed properties publication bounded
and lossless.  `custom_xml::codec::rewrite_props` plans descriptor-based
lexical replacements, computes exact aggregate output length before the one
final `try_reserve_exact`, and avoids intermediate escaped-attribute or
schema-reference buffers.  ID-only edits use a smaller source-offset proof
and do not perform a redundant structural read-back.  Structural edits retain
the configured source event, depth, and XML-attribute-byte checks and preserve
read-back classifications.

MCE branch selection now carries source-offset provenance into the processed
view, so both ID-only and schema-list edits update the selected branch while
leaving inactive branches untouched.  Generated schema references are emitted
from the requested known vocabulary.  They do not clone foreign attributes,
comments, or extension content from an existing reference.  Unknown direct
children of `schemaRefs` remain untouched; removing a reference that carries
opaque attributes, comments, processing instructions, entity references, or
non-whitespace content returns `UnsupportedContent` rather than discarding
those bytes.  Empty-root expansion emits the complete `schemaRefs` wrapper.

The DOCX adapter now propagates malformed internal relationship targets from
all data-store retention callers.  Missing bibliography stores, captured
Custom XML store items, and authored source tags are classified as fixed,
content-free `RefusalKind::NotFound` refusals.

Regression coverage includes:

- `mce_schema_rewrite_maps_the_selected_source_branch`;
- `schema_insertion_does_not_clone_foreign_schema_ref_markup`;
- `schema_insertion_expands_an_empty_root_with_a_schema_refs_wrapper`;
- `schema_edits_preserve_unknown_schema_refs_children_when_shape_changes`;
- `schema_shape_removal_refuses_an_opaque_removed_reference`;
- `lexical_schema_rewrite_admits_exact_output_limit_and_rejects_one_under`;
- DOCX `remove_replace_and_duplicate_tag_flows` refusal assertions.

Only source formatting and repository diff checks were run for this follow-up;
no Cargo, rustdoc, compiler, native, or performance execution was performed.
Independent reviews and the later isolated qualification remain pending.

The 94cc return fixes the source-profile error shape without widening the
policy: depth, XML-event, and XML-attribute-byte admissions now use the
Custom XML `Error::Limit` variants (`Depth`, `XmlEvents`, and
`XmlAttributeBytes`) with the actual observed and configured maximum values.
They are no longer wrapped as generic OPC errors.  The new regression checks
cover event and attribute ceilings alongside the existing depth assertion.

The same return adds an early physical source scan for MCE properties input,
charges each attribute key plus raw value and decoded value before
materialization, and applies the attribute ceiling to payload validation.  MCE
active-offset provenance is now carried in a side table instead of inflating
the temporary XML with one marker comment per candidate; the ID-only route
does not retain that processed XML.  Paired schema-reference frames carry the
reference index (not its byte offset), preserving opaque-content refusal
without a panic.  Exact-source-size MCE rewrite coverage exercises these
bounded routes.
