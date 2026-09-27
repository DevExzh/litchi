# Independent candidate and probe review

Review status: **bounded approval for the coordinator's source, quality, and
paired measurement gates; no static correctness blocker found.** This is a
read-only review. It does not approve production adoption or claim a
performance result.

## Candidate review

The reviewed source is the exact formatted pair under `candidate/applied/`,
compared with `candidate/before/`. The production diff is confined to
`crates/litchi-pptx/src/notes/mod.rs`; the focused tests are test-only in its
`known_uri_tests` module. The implementation keeps the existing private
`resolved` contract and introduces no public API, dependency, cache, global
state, unsafe code, or parser-order change.

For `ResolveResult::Bound(Namespace(value))`, `known_namespace` first checks
the exact byte length and exact bytes of all six existing namespace constants,
then returns that constant's static string. The returned `&'static str` safely
coerces to the input lifetime. Every non-match still reaches the original
`std::str::from_utf8(value).map_err(xml_error)` fallback, so unrelated valid
URIs remain borrowed, invalid UTF-8 remains an XML error, and empty bound
values remain accepted. The `Unbound` and `Unknown` branches are unchanged,
including the existing unknown-prefix error text. Matching is byte-exact and
case-sensitive; no prefix, suffix, length-only, or case-insensitive URI is
accepted as known.

The focused tests independently implement the pre-candidate resolver and
compare both values and typed error text. They cover all six exact constants
and static pointer identity, every replacement byte at every known position
(skipping only the original byte), exact invalid-byte parity,
vendor/prefix/suffix cases including valid Unicode, a Unicode namespace at each
known byte length, empty bound and unbound namespaces, and known and
invalid-byte unknown prefixes. The existing notes scanner differential tests
remain
necessary during the root quality run because this focused module test does
not itself cover the full XML corpus, limits, MCE processing, malformed names,
duplicate attributes, or end-to-end package publication.

The candidate is ADR-compatible on this inspection: the change stays in the
PPTX notes owner, preserves typed refusal behavior and bounded parsing, and
does not alter snapshot, transaction, or public facade layers. The frozen
adoption policy and paired public-workflow evidence remain mandatory.

The archive custody is now coherent: `changed-files.json` names only
`crates/litchi-pptx/src/notes/mod.rs`, and `candidate/files/`,
`candidate/applied/`, and the compatible `candidate/model.patch` describe the
same replacement. The original pre-change source remains retained under
`candidate/before/`.

## Probe call-path review

The slide-only vendor injection does reach this candidate's fallback during
the timed public capture. `Package::opened_presentation` enters
`opened::capture_internal`, which calls
`presentation::capture_slides_with_mce`. Each slide is processed by
`SlidePart::from_part_with_name_with_capture`; its `finish_from_processed`
calls `notes::root_conformance_from_processed` for the slide root. That
scanner runs `scan_processed_xml` and `inspect_element`; for every injected
namespaced attribute, `resolve_attribute` is passed to `notes::resolved`.
The later `notes::load_snapshot_with_slide_root_proofs` call reuses or checks
the same root classifications, so the ordinary initial capture has the
candidate owner in its path.

The generated slide XML has no markup-compatibility namespace. The MCE
preprocessor therefore takes its borrowed no-MCE path and leaves the vendor
declarations and attributes in the bytes scanned by `inspect_element`. Even
when a fixture includes MCE markup, the unknown non-ignorable attributes are
not declared ignorable and remain visible to the scanner. The probe's
post-publication count of every declaration and namespaced attribute, together
with its full text and package readback checks, is the required oracle for the
two vendor controls.

The frozen probe includes only a fixture-side `usize` annotation in
`insert_attributes`, recorded in `probe-build-correction.json`; it changes no
measured path. Its recorded rustfmt check is clean. The candidate remains
archive-only pending the coordinator's baseline, quality, differential, and
paired measurement gates. No Cargo, compiler, native, profiler, or
performance command was run by this reviewer.
