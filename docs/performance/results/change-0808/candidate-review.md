# 0808 independent candidate review

Review status: **bounded approval for build and semantic-gate validation; no
static correctness blocker found.** This review covers the archived source
candidate under `candidate/after/codec.rs` and the exact current-source
control under `candidate/before/codec.rs`. It does not approve production
adoption or make a performance claim.

## Candidate boundary

The before archive matches `crates/litchi-pptx/src/notes/codec.rs` byte for
byte. The after archive has one production path change: the capture-local
`scan_processed_xml` loop uses `quick_xml::Reader<&[u8]>` and a local public
`NamespaceResolver` instead of `NsReader<&[u8]>`. The private inspector takes a
resolver reference. The test-only counter helper remains on `NsReader` and
passes `reader.resolver()` to that inspector. No other file, public signature,
crate dependency, parser configuration, limit, buffered oracle, or unrelated
`NsReader` callsite is changed.

The candidate keeps `NsReader` imported only under `cfg(test)` because the
buffered differential oracle still needs the baseline implementation. That
oracle is therefore an independent control for the direct event transport;
the candidate does not rewrite both sides of the comparison.

## State transitions and error priority

The direct loop matches quick-xml 0.41.0 `NsReader::read_event` and
`process_event`:

1. A pending scope is popped before every underlying read, including when the
   read returns a parser error.
2. A Start or Empty event pushes all namespace declarations before the caller
   performs node/depth accounting or resolves the root and attributes.
3. Empty marks a pending pop after its push, so its own declarations remain
   visible through inspection and disappear before the next event.
4. End marks a pending pop before the existing underflow/decrement branch, so
   the matching scope remains available until the following read.
5. Namespace failures are converted through quick-xml's `Error::Namespace`
   display path before the crate's `xml_error` mapping.

The default resolver retains quick-xml's 256 namespace-declaration cap and
reserved `xml`/`xmlns` rules. Because the push is before the scanner's depth
and node checks, declaration-cap and reserved-prefix errors retain the same
priority as the before path. Root-name resolution, strict/transitional retry,
attribute validation, relationship collection, end-name treatment, missing
root handling, and all byte/node/depth limits remain on their existing paths.

## Focused tests reviewed

The after archive adds these direct-transition checks:

* nested Empty and End scopes with default and relationship rebinding, with a
  later outer relationship attribute as the observable pop result;
* a reserved-prefix error at a depth boundary, proving push-before-depth;
* both reserved-prefix forms and the 256-binding cap, proving cap-before-late
  reserved-prefix priority.

The existing helper and matrix tests retain direct attribute-limit behavior,
Start versus Empty node ceilings, strict and transitional roots, malformed and
duplicate attributes, undeclared prefixes, invalid UTF-8, DTD/PI/CDATA,
missing/unterminated roots, and mutation/corpus parity. Every comparison
includes the unchanged buffered `NsReader` oracle's structured result or
formatted refusal.

## Findings and conditions

I found no static semantic or ownership blocker in the archived diff. The
candidate can proceed through the coordinator's ordered workflow: archive
format/source checks and the before-build/probe-quality gates first, then fresh
application, production quality gates, after-build checks, and the frozen
paired native/allocation matrix and owner-scoped profile. The numerical and
semantic artifacts must be replayed independently before the public adoption
gate is evaluated. A passing static review is not evidence of speedup, lower
allocations, or production eligibility.

The candidate remains private to the notes scanner and does not justify a
cross-format claim. If any fresh build or differential oracle reports a
changed refusal string, first-refusal position, root dialect result,
relationship inventory, or output/semantic value, this review is withdrawn in
favor of that evidence.
