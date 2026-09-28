# 0808 source review: direct notes scanner event handling

This is a read-only source review of the archived 0808 candidate at the
current base `d28e3dc702`. It reviews the candidate against the 0807
event-transport follow-up and the accepted constraints in ADR 0001 (API and
ownership layers), ADR 0003 (immutable snapshots and atomic validation), ADR
0005 (bounded input and measured performance), ADR 0006 (preservation,
validation, and typed refusal), ADR 0008 (verification), ADR 0013 (PPTX notes
ownership), and ADR 0032 (snapshot-derived values). No production source was
applied and no Cargo, formatter, workload, or profiler command was run by the
reviewer.

## Scope and source identity

The current production file and the archived candidate `candidate/before/`
file are byte-identical (SHA-256
`8485d43c99f19bda9b3510c8323aa6bc2b117fb7df262372ce82868a5397239c`). The
candidate `candidate/after/codec.rs` changes only the notes codec. Its
production changes are limited to `scan_processed_xml` and the private
`inspect_element` resolver argument; its test-only counter helper remains on
`NsReader` and passes `reader.resolver()` to the inspector. The independent
buffered scanner oracle remains a `NsReader` implementation and is unchanged.

The candidate removes only the intermediate `NsReader` event handoff from the
capture-local scanner. It does not change `process_ooxml`, parser configuration,
the buffered scanner, the public notes API, another `NsReader` caller, any
crate dependency, or any production limit constant. `Reader::from_reader` is
still the slice reader over the already-processed bytes, and
`trim_text(false)` is retained.

## Event-state equivalence

The archived before path uses quick-xml 0.41.0 `NsReader::read_event`. That
method pops a pending scope before reading, then `process_event` pushes a
scope for `Start` and `Empty`, marks `Empty` and `End` for the next pop, and
passes non-element events and parser errors through. The candidate spells out
the same state machine:

* `pending_pop` is consumed at the top of every loop, before
  `Reader::read_event()` and before mapping a parser error;
* `Start` calls `NamespaceResolver::push` before node/depth accounting and
  before `inspect_element`;
* `Empty` calls `push`, marks `pending_pop`, then performs the existing node,
  depth, and inspector branches;
* `End` marks `pending_pop` before the existing depth-underflow check and
  decrement; and
* text, comments, declarations, CDATA, EOF, and parser errors do not modify
  namespace state beyond the pending pop already handled at the loop boundary.

This preserves the important Empty rule: its declarations are visible while
the element is inspected, but disappear before the following event. An End
scope likewise remains in force through processing of the returned End event
and is removed before the next read. If that next read returns a parser error,
the scope has already been removed, matching `NsReader::read_event`.

`NamespaceResolver::push` is called with its default per-element declaration
limit, so the candidate retains quick-xml's 256-declaration cap and the same
reserved-prefix checks. Mapping `NamespaceError` through
`quick_xml::Error::from` before `xml_error` produces the same display path as
the `?` conversion in `NsReader::process_event`. The candidate therefore
retains the namespace error's priority over the scanner's later node/depth
checks and over checked-attribute inspection.

## Scanner semantics

`inspect_element` now receives `&NamespaceResolver` directly and calls the
same `resolve_element` and `resolve_attribute` methods. The test-only helper
continues to obtain that resolver from its `NsReader`; it does not alter the
production transport under test. The root namespace and
local-name checks, relationship inventory, attribute duplicate/malformed
checks, unescape behavior, attribute counters, and all result ownership remain
unchanged. End-tag names were not checked by the before scanner or by this
candidate; both continue to rely on the same quick-xml reader configuration
and the existing depth/missing-root checks. This is required for exact parity
with the current refusal behavior.

The root retry still tries Transitional and then Strict, masking a pair of
failed scans to the generic root error. The candidate keeps the raw and
processed byte ceilings, node and depth ceilings, DTD/PI/CDATA refusal,
unexpected-close handling, and missing/unterminated-root result in the same
places. Namespace declarations are installed before the first root check, so
both default and prefixed roots retain their existing resolution behavior.

## Focused semantic coverage

The candidate adds focused differential cases for nested Empty elements with
rebound default and relationship namespaces, namespace-error priority over a
depth limit, both reserved prefixes, and the declaration-cap-before-reserved-
error boundary. The existing direct-vs-buffered tests retain both
Transitional and Strict inputs, Empty and Start paths, root and non-root
resolution, malformed and duplicate attributes, unknown prefixes, invalid
UTF-8, DTD/PI/CDATA, missing roots, depth/node limits, and independent raw and
processed ceilings. The corpus lane compares up to 600 deterministic PPTX XML
parts and structural mutations against the unchanged buffered oracle.

The nested-namespace test also checks that a relationship attribute after
several Empty and End transitions resolves through the outer binding. The
depth-priority test places a reserved `xml` declaration at the first level
past the depth boundary; its expected error proves that `push` precedes the
scanner budget check. The declaration test places the reserved declaration
after 256 accepted namespace bindings and expects quick-xml's declaration-cap
error first.

## Review result

**Bounded approval for candidate build, focused semantic gates, and the
planned matched workflow measurement; no static correctness blocker found.**
The source preserves the selected 0807 seam and its state/error ordering. This
review does not approve production adoption, does not establish a latency or
allocation result, and does not replace the coordinator's required quality,
independent replay, profile, and public workflow gates. Any compiler, oracle,
refusal, or measurement discrepancy supersedes this static review.
