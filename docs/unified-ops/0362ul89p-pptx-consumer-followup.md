# 0362ul89p — PPTX keep-going integration-consumer follow-up

Date: 2026-10-25. This is a source-only repair of integration-test consumers;
it does not claim a compiler, runtime, performance, or qualification result.

The frozen source base is
`15d8d53a0cf6d27f350b866463b4235bf26b08d9`. The keep-going diagnostic log is
outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-15d8d53a0c/01-pptx-all-features-check-attempt-actual-f5-01.log`
(100,739 bytes, SHA-256
`a31129804840cddda84361dc4201852856d55afc25de7f0d49cc7d5868295084`). Its
diagnostic index is outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-15d8d53a0c/diagnostics-index.json`
(18,586 bytes, SHA-256
`c25bee0ed0c486cd54e2d132c928509bfb08484991437136c53c403104b5d3be`). The
index contains 19 diagnostics in six distinct test files, beyond the limited
first-target evidence recorded by 0362ul89o:

* `shape_chunk_publication.rs`: two expression-valued `Error::Limit` patterns
  now bind `observed` and `maximum` and guard their exact expected values;
  the typed `CodecLimitKind` matches and both historical replacement constants
  remain asserted.
* `pptx_readonly_recommended.rs`: source-limit qualification uses the imported
  type, and the event-count helper uses fallible part iteration and the real
  `Part` payload rather than treating `PartMetadata` as a payload owner.
* `slide_insertion.rs`: the duplicate `Error::Limit` alternative is removed;
  the limit-or-OPC refusal assertion remains unchanged in meaning.
* `pptx_tolerance_report.rs`: the animation owner `Limits` profile is imported
  for every explicit fixture boundary.
* `pptx_metadata_families.rs`: the structure owner limit is imported under a
  disambiguating name, retaining the explicit default profile beside the
  guides profile.
* `source_related.rs`: eager notes text and eager legacy comments are read
  with their required explicit owner profiles; source-side assertions and
  exact values are unchanged.

The static sibling audit covered all 108 tracked PPTX integration-test Rust
files and all six examples in this sparse source tree. Required raw
parser/writer calls retain explicit profiles, including animations, comments,
media, notes, tags, laser, model3d, shape, transition, presentation
properties, and metadata families. No production source, public API, fixture,
default fallback, or return shape was changed. The task wording says seven
test files, but the preserved index has six distinct test paths; all 19 indexed
test diagnostics are addressed here.

No Cargo, rustdoc, build, runtime, or performance command was run. The source
base and external diagnostic identities above are the only validation evidence
for this record; repository diff, evidence-size, binary, and path guards are
required before commit.
