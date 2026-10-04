# 0896 source and methodology review

This review covers the minimal QName expansion candidate in the working tree
against committed production `3ff493b638`. It is a source review; it does not
admit a performance result or production adoption.

## Source finding

`crates/litchi-ooxml-common/src/mce/stream.rs:3352-3384` now validates the
lexical value with the existing `xml_name::is_qualified_name` predicate and
borrows the prefix and local slices with `split_once(':')`. The removed path
constructed an owned `QualifiedName`, then immediately borrowed those same
slices. For every valid input, the two paths therefore pass the same XML 1.0
NCName/QName grammar and select the same prefix and local parts. For every
invalid input, both return the stream's existing typed `invalid QName`
refusal.

The namespace and limit order remains unchanged. `parse_element` still checks
the complete lexical name before expansion, installs the declarations for the
start tag before resolving names, applies the default namespace only to an
element, leaves an unprefixed attribute in `NO_NAMESPACE`, and delegates the
fixed `xml` binding to `Scope::resolve`. URI, local-name, unbound-prefix, and
duplicate-attribute checks remain after the same expansion points. This keeps
the allocation reduction inside the MCE stream owner and preserves the
callback, recovery, typed-error, and bounded-resource contracts required by
ADRs 0005 and 0006.

The grammar equivalence follows directly from
`crates/litchi-ooxml-common/src/xml_name/codec.rs:8-25` and
`crates/litchi-ooxml-common/src/xml_name/validation.rs:25-32`: after the
predicate succeeds, `QualifiedName::prefix()`/`local()` return exactly the
same slices as the candidate's `split_once` result. No normalization or
namespace lookup is performed by either implementation.

## Test review

The added stream tests cover default versus unprefixed-attribute namespaces,
prefixed and Unicode QNames, malformed colon structure, an overlong lexical
name, and unbound-prefix/error precedence. Existing stream suites additionally
cover namespace rebinding, default resets, aliases, duplicate expanded names,
hidden branches, recovery, callback ordering, and exact stream limits.

The valid-name test uses `xml_name::parse` as its parser reference. Its name
refers to that shared QName parser; it does not need to invoke the legacy MCE
processor because this candidate only replaces construction of the already
validated QName before namespace expansion. The broader differential receipt
still covers the stream and legacy MCE processors independently.

## Measurement scope

The direct signal should come from the existing `xml_attribute_bounds`
`mce_stream_*` cases, especially the benign worksheet/document, counted
worksheet/document, review-long-name, and long-element cases. The producer
full-cell scenario remains necessary as an end-to-end guard, but its complete
producer worksheet reaches `<sheetViews>` and the selected scanner refuses
there (`tools/perf-baseline/src/producer_shape/selected_route_tests.rs:269-286`).
The current `SourceWorksheet` path then falls back to the authoritative stored
parser (`crates/litchi-xlsx/src/workbook/source.rs:1034-1068,1205-1213`), so
this candidate only affects the prefix scanned before that refusal in this
corpus. The rejected 0894 metadata route must not be used to claim current
production benefit.

The eventual report should separate the isolated stream effect from whole
XLSX CRUD results, retain DOCX/PPTX and non-stream XLSX controls for shared
process noise, and make no adoption claim until the current-source quality,
semantic differential, resource/error checks, and matched native measurement
all pass.

## Final source-review receipt

Review base: committed source `3ff493b638`.

The final candidate files reviewed here have these SHA-256 identities:

* `crates/litchi-ooxml-common/src/mce/stream.rs`: `597a65d9070f65c48669a592dd2715fc50e3461936ce9a6e65d0779e2d79464e`
* `crates/litchi-ooxml-common/src/mce/tests.rs`: `ed0f3f5c748b58d7184068e8de29bb2b05e48d7c73f25718f580ef8b5aa15682`

The earlier formatted test-only candidate (`89605bb3e89d691bbe519564cf62a3e5bcc5acb9fa296864452e5ec7b31a1591`) is retained in the packet's relocation receipt; the final source uses the equivalent literal string required by Clippy.

The candidate diff is formatted and passes `git diff --check`. Source review
status: **approved for the 0896 quality, differential, resource/error, and
matched native-measurement gates**. This is a source-level approval only; it
does not admit the candidate for production adoption or claim a whole-producer
speedup. The frozen primary producer gate must decide that separately.
