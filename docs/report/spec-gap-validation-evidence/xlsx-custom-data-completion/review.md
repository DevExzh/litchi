# XLSX Custom Data completion — independent final review

**Final disposition: PASS.** This bounded review found no concrete blocker in
the six previously raised dispositions for the corrected frozen source. It does
not widen the completion contract or claim native Office acceptance.

## Frozen source and validation

- `gates/freeze.json` names `/tmp/litchi-custom-data-completion` at base
  `1d39eea516248c758b8ea429c879a4a3245e4fbb`. Recomputed SHA-256 values match
  all ten selected Rust files and `Cargo.lock`, including the corrected
  `embedded_data.rs` value
  `0d0a8407724d4ba4c53e2613b547ff227268ac88f8b3e3003bd5d6038663a815`.
  The isolated checkout is the authority for this receipt; the shared
  workspace has a different ignored `Cargo.lock` and is not used to override the
  freeze.
- The isolated gates report 1,860 XLSX, 619 OPC (one pre-existing ignored
  external ZIP64 case), and 302 common tests, with zero failures. Strict
  Clippy, rustdoc, boundaries, diff checks, and the selected-file formatting
  batch pass. The only whole-crate formatting failure is the unchanged
  baseline `drawing_svg_read.rs` exception recorded by the gate runner.
  Sources are stable before and after the run; see
  [freeze](gates/freeze.json), [results](gates/results.json), and
  [verification](gates/verification.json).
- The corrected [edge capture](../../../../crates/litchi-xlsx/src/connections/embedded_data.rs#L482)
  checks the existing relationship limit and then fallibly reserves one edge
  at a time. The earlier diagnostic measured 8.25× baseline requested bytes on
  the small read and 82× baseline peak live bytes on small remove-inverse,
  including about 6.5 MiB reserved for one edge by the broad pre-reserve.
  The corrected source removes that allocation amplification without changing
  relationship semantics. The final candidate profile is recorded separately
  by the profiling receipt.

## Six dispositions

1. **PASS — zero events and `Empty` depth.**
   [`preflight_xml_with_limits`](../../../../crates/litchi-xlsx/src/connections/codec.rs#L279)
   runs before PI copying, MCE processing, and DOM/NsReader work. Its event
   budget is passed unchanged, so a zero budget rejects the first event; both
   `Start` and `Empty` events count nodes and prospective depth. The query-table
   path uses the same preflight. Exact zero-event and empty-element depth tests
   pass without widening the profile.

2. **PASS — lexical preflight, PI copying, and MCE output.**
   The borrowed lexical preflight precedes
   [`strip_processing_instructions_with_limit`](../../../../crates/litchi-xlsx/src/connections/codec.rs#L564)
   and MCE processing. PI-free input stays borrowed; an owned PI result is
   bounded by the temporary ceiling. The
   [`contains_mce_markup`](../../../../crates/litchi-xlsx/src/connections/codec.rs#L417)
   byte test uses the same MCE namespace window as the common processor’s
   borrowed-versus-owned branch. Consequently MCE output is capped by
   `min(max_temporary_bytes, max_xml_bytes)` only on the owned path; a borrowed
   input may be larger than the temporary ceiling while remaining under its
   input ceiling.

3. **PASS — namespace, prefix, name, and CDATA bounds.**
   Attribute counts and bytes, names and prefixes, namespace bytes, text, and
   CDATA are checked before `NsReader` processing. Resolved namespace URIs
   are entity-decoded before recognizing connection bindings, so escaped URIs do
   not bypass or miss a binding. The namespace alias, URI-entity, prefix,
   attribute, and CDATA boundary cases pass.

4. **PASS — retained source versus replacement output.**
   The common codec’s `*_with_limits_and_output` functions validate retained
   input independently and preflight changed output before replacement
   allocation; semantic no-ops retain the source `Arc`. Connection publication
   applies the full XML cap, including `min(max_connections_bytes,
   max_output_bytes, max_temporary_bytes)`, while inverse restoration may retain
   an admitted source above the temporary ceiling. Content-type transitions use
   the existing OPC `plan_edit` API. Shrink, growth-refusal, moved-payload,
   and inverse tests cover these rules.

5. **PASS — candidate aggregate output accounting.**
   Snapshot and staged-output preflights sum modeled parts, content types,
   relationship XML, and each connection’s own relationship token, including
   its canonical empty relationship view. The commit path checks the candidate
   aggregate before publication and validates it again after applying edits.
   The accounted output is logical uncompressed modeled OPC content, not ZIP
   archive size.

6. **PASS — relationship provenance and source checking.**
   `Bindings::Source` retains `OwnedRelationships`, including physical-member
   presence. Source comparison therefore includes relationship presence and
   bytes along with the retained source XML and values. Restore replaces the
   XML proof and restores the source relationship provenance. Tests cover
   absent versus explicit-empty relationship members, stale source refusal, and
   exact inverse restoration.

## Scope and limits

The temporary ceiling bounds individual XML replacement and PI/MCE
transformation buffers; it is not an operation-wide allocator budget. Retained
source storage is admitted under input limits and may exceed that temporary
ceiling. Output accounting covers modeled parts, content types, and relationship
tokens, including canonical empty relationships; it is not compressed ZIP size.
Custom Data remains inert, and native producer/Office acceptance is outside this
receipt. Existing unsupported graph/MCE shapes and changed signed packages remain
refusals. This review made no source edits, broad test reruns, or commit.
