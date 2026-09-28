# Change 0823 scanner candidate

This directory archives a candidate against commit `b76786208d`. It is not a
landed implementation and carries no performance claim.

The candidate removes the production `NsReader` wrapper from
`shape::reader::Scanner::scan`. The scanner constructs
`Reader<&[u8]>` and owns one `NamespaceResolver`, then resolves only the
element name needed by each `Start`, `Empty`, or `End` arm. Text, CDATA,
general references, comments, declarations, and processing instructions keep
the existing event handling. The old `NsReader` loop is retained under
`cfg(test)` as `scan_with_nsreader_oracle` so the candidate can be compared
against the pre-change implementation without a parser counter or a second
production route.

The manual resolver follows quick-xml 0.41.0's pinned
`NsReader::read_event` sequence:

1. Capture the pre-event position and decoder.
2. Pop the deferred `Empty`/`End` scope immediately before the underlying
   reader call.
3. Read one borrowed event.
4. Push declarations for `Start`/`Empty`, set the deferred pop for
   `Empty`/`End`, and convert `NamespaceError` with
   `map_err(quick_xml::Error::from)?`.
5. Capture the post-event position.
6. Resolve the element name with `resolve_element` and run the unchanged
   scanner arm.

The `End` arm resolves before the deferred pop, which is required when a
prefix is shadowed on the opening element. The push occurs before the
post-event position, preserving the original resolver-error precedence. The
candidate does not clone events or namespace state.

The differential tests compare the complete private record vector and retained
string arena (using their debug representation) on success, and the typed
error debug representation on failure. The cases cover nested and shadowed
bindings, default namespace reset, an empty-element scope, end-element
resolution, a UTF-8 BOM, default and Strict roots, an unknown root prefix,
reserved `xml` and `xmlns` declarations on both Start and Empty events, PI and
DTD refusal, unclosed markup, CDATA, entities, duplicate attributes, the exact
256/257 declaration ceiling in both Empty and Start forms, and scanner limits
for nodes, depth, shapes, and retained text. The existing public Scene/MCE
tests remain the input/output-boundary coverage; this archive does not add a
second prelude oracle for those limits.

The candidate uses no retained memo, cache, global state, or unsafe code, and
introduces no per-event allocation. The resolver's existing bounded namespace
storage remains the same kind of temporary state that `NsReader` already
owned. The candidate changes no public API and does not alter the source bytes
or span accounting. It therefore has no ADR 0030 or ADR 0032 surface: it is a
local parser-work reduction over the same borrowed payload and temporary
resolver. The existing MCE and public Scene tests remain the required follow-
up because this archive's oracle exercises the scanner after MCE preprocessing.

Archive integrity:

* `before/reader.rs` is byte-identical to
  `crates/litchi-pptx/src/shape/reader.rs` at the base:
  `58e88e86ffc9d4053b694d7928377a05beacbfcb47368cefc9d2fc229b1af26e`.
* `after/reader.rs` SHA-256:
  `a08a48f25c9e2e5f367bd67c579614f1c0c1311784ec81dd30d201b69c3db7f2`.
* `candidate.patch` is the unified diff between those two source files:
  `e4a97bd62ec14043922ba0babeb08bcf00a0e6ce110205d4ff67d7d14cd347c3`.

No Cargo, compiler, formatter, or benchmark command was run by the candidate
agent; those executions remain coordinator-owned.
