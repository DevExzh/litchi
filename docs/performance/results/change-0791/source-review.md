# 0791 — current PPTX notes-scanner source review

This is a bounded source review for the post-0785 PPTX capture path. It does
not change production code, approve an optimization, or claim a speedup. The
current 0791 probe reuses the 0784 operation boundary and fixture while the
known-namespace-URI change from 0785 is present in production. Its current
profile and paired native controls are the authority for deciding whether
either hypothesis below is worth implementing.

## Evidence and scope

The 0783 phase capture made up 69.051% of the large generated lifecycle
median; serialization made up 23.320%, and commit, staging, and apply made up
4.290%, 2.671%, and 0.631%. That makes capture-local work a reasonable place
to look, but these are phase shares of one warm, in-memory borrowed-package
fixture. They do not describe source-backed, cold-cache, range-source,
native-Office, or concurrent workloads.

The 0784 large Callgrind records place the selected per-slide edge of the
inclusive notes scanner at `scan_processed_xml` (374,786,958 guest
instructions) and its `inspect_element` child at 204,004,638. The aggregate
`scan_processed_xml` value across all incoming edges in that profile is
377,129,631. These are different scopes and must not be compared as if they
were one measurement. The largest scanner-side self rows include
`inspect_element` (34,965,460), the quick-xml attribute iterator
(`IterState::next`, 38,981,486), and quick-xml event parsing. These are
diagnostic guest instructions from the pre-0785 profile; they are not native
cycle shares. The 0785 exact namespace comparison is already integrated and
reported an 8.489% paired large-capture p50 improvement and a 6.844% paired
large lifecycle improvement on that fixture. Those results cannot be added to
a future result or used to predict one.

## Primary hypothesis: skip the empty attribute iterator

The narrow seam is
[`inspect_element`](../../../../crates/litchi-pptx/src/notes/codec.rs:447)
after the existing element namespace, local-name, and root checks and before
the `checked_attributes()` loop at line 474. A candidate diagnostic change is:

```text
if element.attributes_raw().is_empty() {
    return Ok(());
}
```

quick-xml defines `BytesStart::attributes_raw()` as the bytes after the parsed
element name. Its reader removes the `<` and, for an empty element, the final
`/>` before constructing `BytesStart`; therefore an exact empty slice means
there is no attribute or attribute syntax to inspect. The `CheckedAttributes`
constructor is allocation-free for ordinary small tags, but its first
`next()` still constructs and advances the quick-xml iterator. The branch can
remove that known-empty call while leaving all element validation in place.

The branch must be exact. A tail containing one or more spaces, tabs, or
newlines is nonempty and must continue through `checked_attributes()`: malformed
attribute syntax must still produce the existing error. Namespace declarations
also make the raw tail nonempty, so the branch cannot bypass namespace binding
or validation. No attribute means there is no duplicate name, attribute-value
UTF-8, entity, relationship, attribute-count, or decoded-attribute-byte check
to perform. The branch consequently does not change any of those checks for an
input on which they can arise.

This hypothesis has a bounded ceiling. It can remove only the empty-tag
iterator path; it does not remove quick-xml event parsing, namespace
resolution, element-local UTF-8 validation, or the scanner's XML limits. The
retained 0784 records do not break the iterator calls down by zero-attribute
versus attribute-bearing tag, so even the fraction of the roughly
53-million-instruction `IterState::next` inclusive row that is removable is
unknown. The current 0791 Callgrind result should compare the exact iterator
symbols and total capture instructions before any source change is considered.

The focused hardening matrix for this candidate must include both `Start` and
`Empty` events with no attributes (`<p:x></p:x>` and `<p:x/>`), spacing-only
tails (`<p:x />`, tabs, and newlines), namespace declarations, malformed
attribute syntax, duplicate names, invalid attribute names and values, bad
entities, and the XML node/depth/attribute limits. The existing
`buffered_scan_oracle` in `notes/codec.rs` and its handcrafted, mutated, and
repository-corpus differential tests compare both projected relationship IDs
and typed refusal text; they are the right oracle and should remain unchanged.
No `checked_attributes` bypass, unchecked iterator, or fast return may be
introduced before the existing namespace/local/root validation.

## Secondary hypothesis: return the successful presentation scan

There is a separate, mechanically redundant pair in
[`load_index_with_slide_root_proofs`](../../../../crates/litchi-pptx/src/notes/package.rs:426):
it first calls `root_conformance` at lines 435–435 and then calls `scan_xml`
again at lines 436–440 with the conformance just found. `root_conformance`
already scans the complete processed XML once for Transitional and, when
needed, once for Strict. A private helper could return
`(Conformance, XmlScan)` from the first successful `scan_xml`; the existing
`root_conformance` wrapper could keep returning only the conformance for its
other callers. `load_index_with_slide_root_proofs` would consume the returned
scan directly.

This preserves the current retry and generic-error masking: Transitional is
attempted first, Strict is attempted only after a Transitional refusal, and a
failure from both attempts still becomes `invalid {root} root or namespace`.
The successful scan's `XmlScan` is exactly the value the second call currently
produces. It does not change slide-root proofs, MCE processing, relationship
validation, graph ownership, or publication behavior.

This is lower priority because the 0784 dominant scanner call path is the
per-slide `SlidePart::finish_from_processed` route, while the duplicated scan
is over the presentation root. A fresh profile or a small diagnostic counter
must establish its call count and inclusive cost on the current source before
it is treated as useful. The test oracle should exercise valid Transitional and
Strict roots, an invalid root, raw and processed size limits, and failures in
each attempted conformance, comparing the old two-call result to the helper's
result and error masking. Do not merge this with a broader scanner rewrite.

## Measurement required before a production experiment

For either hypothesis, use an archive-only candidate against the current
production revision and the existing 0791 owner-scoped probe. Run the normal
quality gates and the full notes differential oracle first. Then repeat the
same two Callgrind passes and six alternating native blocks (tiny, medium, and
large; three warmups and thirty samples; CPU 12) used by the frozen 0791 plan.
Retain p50, p95, p99, total guest instructions, the relevant iterator/scanner
call counts, allocations, peak operation bytes, RSS, source/output identities,
and semantic readback. A candidate must show a reproducible operation-level
benefit without a latency, allocation, memory, error-order, or identity
regression. Callgrind is useful for local work removal but is not a native
latency claim; frame-pointer samples are only qualified when the exact capture
owner remains on the stack.

The 0784 fixture is deliberately shape-heavy and generated. A passing result
would still require controls with attributes, namespace declarations, vendor
namespaces, and malformed inputs before adoption. No retained cache, new
execution model, unsafe code, public API, or architecture change is needed for
either local hypothesis. The broader CRUD goal remains open.

## Current 0791 ranking

The current frame-pointer records strengthen the ranking only as an attribution
signal. Repeat 0 has 1,066 qualified owner samples out of 3,143 total, and
repeat 1 has 1,068 out of 3,153; one and two samples, respectively, contain an
unknown interior frame. Counts below are nested, include warmups, and overlap:

| Nested symbol | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| `scan_processed_xml` | 935 | 929 |
| `inspect_element` | 318 | 289 |
| `CheckedAttributes::next` | 186 | 148 |
| `IterState::next` | 134 | 120 |
| `resolved` | 5 | 4 |
| `load_index_with_slide_root_proofs` | 8 | 5 |

These counts do not form phase fractions and do not prove a speedup. The
current large Callgrind rows give a more direct local-work bound: exactly
181,678 `inspect_element` calls and 273,263 `CheckedAttributes::next` calls
arrive from that inspector (273,265 from all incoming edges). The corresponding
`scan_processed_xml` row has 104 incoming calls in the retained large profile.
The current records therefore keep the empty-tail branch as the first
candidate, while the presentation-scan fusion remains lower priority. The
profile does not report how many of the 181,678 tags have an exactly empty
attribute tail, so it still cannot predict the branch's savings.

The quick-xml registry source also shows that its attribute iterator tracks
seen names in a `Vec<Range<usize>>` and switches to a hash prefilter after its
small-tag threshold. The local `CheckedAttributes` wrapper has its own bounded
quick-xml phase and a later ordered-map phase. This is a possible later
duplicate-check work hypothesis, but the current evidence contains no
allocation count for that internal state and makes no allocation claim. It
must not be turned into an unchecked-attribute path or implementation change
without a separate duplicate-error-order, hostile-tag, allocation, and paired
latency study.
