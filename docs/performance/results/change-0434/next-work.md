# Batch 0434 next work: bounded ODT authoring

This is a source and coverage audit, not a performance result. The 0434 change
only groups safe ODS text bytes while preserving scalar-limit and cancellation
semantics. The next implementation should address the separate ODT/ODP fresh
authoring gap identified in the [goal audit](../../GOAL_AUDIT.md), without
folding it into the current ABBA comparison.

## 1. ODT is the smallest next creation slice

The current ODT builder retains the semantic model in many vectors, including
`elements` ([`builder/model.rs`](../../../../crates/litchi-odt/src/builder/model.rs#L47-L67)),
then constructs a complete `content.xml` body in one `String`
([`builder/codec.rs`](../../../../crates/litchi-odt/src/builder/codec.rs#L7-L22))
and passes that complete buffer to the package writer
([`builder/package.rs`](../../../../crates/litchi-odt/src/builder/package.rs#L31-L49)).
The existing `OdtSemanticCreateSmall` harness case measures only this tiny
whole-buffer path ([`perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs#L31583-L31595));
the ODT text-to-sink case serializes an already-open document and is not
authoring evidence.

The bounded follow-up should be an explicitly narrow, opt-in ODT paragraph
writer built on the existing common generated-XML publication seam. Its first
contract should be:

- consume an ordered paragraph source once;
- emit only ordinary `text:p` fragments through a reusable bounded fragment
  window;
- retain standard ODT envelope, manifest, deterministic serialization, XML
  auditing, cancellation, row/text/XML/output limits, and typed sink failure;
- exclude tables, lists, forms, fields, rich runs, media, hyperlinks, unknown
  source content, and existing-document append until separately designed.

The proof gate should reopen the output with `litchi-odt::Document`, compare
paragraph/text semantics and exact archive/XML hashes, check compact package
topology, and exercise one-under/equal/over limits, invalid XML text,
cancellation, short/partial sinks, producer failure, and zero-output
preflight. Timing must include the fresh producer and sequential publication;
reopen/oracle work stays outside it. Allocator and process-RSS observations
remain separate descriptive evidence. Keep the selector opt-in until an
independent baseline/candidate ABBA capture passes all identity gates.

## 2. ODP follows after the ODT seam

ODP has the same missing bounded creation proof but a larger first surface:
`Builder` retains `Vec<Slide>` and related optional collections
([`authoring/builder.rs`](../../../../crates/litchi-odp/src/authoring/builder.rs#L56-L64)),
builds the complete slide body and `content.xml`
([`builder.rs`](../../../../crates/litchi-odp/src/authoring/builder.rs#L892-L927)),
and the current `OdpSemanticCreateSmall` case measures only tiny buffered
creation ([`perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs#L34565-L34577)).

After the ODT provider and common seam are independently reviewed, use a
one-slide/plain-title-and-text ODP stream. Require the
`office:body/office:presentation` catalog shape, deterministic package and
compact XML, `Presentation` readback, bounded fragment and text limits,
cancellation, and sequential sink failure evidence. Keep shapes, media,
animations, notes, declarations, and rich slide layout outside the first
selector; those features need their own bounded grammar and preservation
evidence.

## 3. Native compatibility stays a separate gate

Native evidence is narrow but real: the retained ODT and ODP LibreOffice
resave/readback chains and real-producer corpus tests establish correctness for
selected fixtures. They do not establish broad producer compatibility, bounded
authoring, or performance. The next compatibility check should rerun the
locked native-resave chain after its existing lockfile gate is resolved, then
add at least one additional ODT and ODP producer/shape or explicitly retain the
one-fixture-per-family limitation. Keep this gate separate from the ODT/ODP
streaming timing selector.

## Remaining distinctions

Fresh creation does not cover logical append to an existing document, package
Part addition, or arbitrary modification/repackaging. Existing ODT/ODP edits
and native readback should not be counted as forward-only append evidence.
Those scenarios remain later, separately measured obligations under the
[CRUD checklist](../../../../docs/CRUD_Scenario_Checklist.md) and accepted ADR
tree. No change in this plan closes the broader non-iWork goal.
