# Batch 0435 next work

The 0435 candidate adds a bounded fresh plain-text ODT path and typed common
publication/diagnostic seams. It does not close the broader non-iWork goal, and
this note records the measured memory/CPU tradeoff without a general speedup claim.

## Completed ODT fresh-creation evidence

Use the existing full-buffer builder as the semantic control and the new
streaming selector as the candidate. The old builder constructs a complete
body/content XML and archive vector; the new path consumes ordered paragraphs
once and publishes reusable fragments through the common writer. Relevant
source is [`streaming.rs`](../../../../crates/litchi-odt/src/streaming.rs#L1),
the common publication methods in
[`writer.rs`](../../../../crates/litchi-odf-common/src/core/writer.rs#L1210),
and the typed limit path in
[`generated_xml.rs`](../../../../crates/litchi-odf-common/src/core/generated_xml.rs#L23).

The passing formal matrix binds the source,
build, corpus, semantic reopen, five-member topology, fixed styles/meta
identities, and output sinks. It retains 36 reports, 1,080 samples, six profiles,
and every regression flag. The operation timer covers paragraph
generation and publication, while corpus setup, reopen/oracle work, process
probes, and final diagnostics stay outside it. Latency, allocation counters,
RSS, and any scaling result need separate treatment; a bounded fragment window
does not prove a whole-process memory bound. Failure, cancellation, exact-limit,
fallible-source, short-sink, and typed XML-attribution checks remain part of the
acceptance boundary.

The ODT slice is fresh plaintext authoring only. It does not measure or imply
logical append, adding a package Part, arbitrary existing-document edits, or
repackaging. Rich ODT semantics and external native producer validation remain
separate work.

## First investigate ODT Work charging

Streaming large-case allocator peak is 420,091 bytes versus 22,450,985 buffered,
with normal p50 at 3.351 / 3.369 times candidate Builder latency. The streaming
whole-process profile assigns 45.74% self samples to ExecutionContext::consume.
Test bounded ordinary-text Work batching while preserving per-scalar cancellation
and exact refusal fallback. Use committed 0435 streaming as the measured baseline;
do not infer a causal speedup from the profile alone.

## Then follow with ODP fresh creation

After the ODT CPU tradeoff review, the next narrow authoring target
is an ODP fresh package with one family-owned plain title/body or scalar slide
grammar. The existing ODP builder is buffered in
[`authoring/builder.rs`](../../../../crates/litchi-odp/src/authoring/builder.rs#L56);
the common XML/archive seam can be reused, but ODP namespaces, slide grammar,
and semantic oracle must remain ODP-owned. Start with a fixed minimal topology;
defer rich shapes, media, animations, append, Part addition, and repackaging.

## Keep the remaining evidence boundaries explicit

The [goal](../../../../docs/GOAL.md) and
[CRUD checklist](../../../../docs/CRUD_Scenario_Checklist.md) still require
independent coverage for fresh creation, logical append, package-Part addition,
and arbitrary modification/repackaging. Cold/range I/O, native producer
coverage, allocator/RSS/scaling evidence, and richer ODF authoring remain open
even when fixture correctness is green. The native runtime probe is unavailable
in this batch; fixture tests should be retained as a separate evidence class,
with later LibreOffice/Office runtime checks required before claiming native
compatibility.

