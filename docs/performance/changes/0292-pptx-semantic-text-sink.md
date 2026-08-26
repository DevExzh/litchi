# Change 0292: PPTX semantic text sink

Date: 2026-08-27

Status: Accepted bounded streaming evidence

`performance_claim: none`

## Decision

The `litchi-pptx` owner exposes `write_text_to` on both the borrowed eager
`Presentation` view and `SourceBackedPresentation`. Each traversal submits one
`TextObjectKind::Slide` to the caller-owned `SequentialTextWriter` for each
logical slide. `TextOutputOptions` controls empty-object inclusion, paragraph
and slide separators, and output/object ceilings. The returned
`TextOutputReport` records the exact accepted byte and object progress.

The sink path does not construct the document-wide `String` returned by
`text()` and does not retain an all-slide semantic text `Vec`. At any point it
retains only the current slide's bounded raw/processed/text parse state. Empty
slides are either emitted or omitted according to `include_empty_objects`; the
writer, rather than the PPTX traversal, owns separator and limit accounting.

## Eager package path

The eager path performs a complete relationship preflight before writing the
first sink byte. Its `SlideReference` collection contains presentation-order
relationship metadata only; preflight does not parse or materialize slide
payloads. It then resolves and parses one slide at a time and passes that
slide's semantic text to the writer.

This does not change eager OPC ownership. An eager `OpcPackage` continues to
retain the admitted package Parts, including slide payloads. The bounded sink
claim applies to semantic output accumulation, not to the package's existing
input retention.

## Source-backed lazy path

The source-backed path reuses the validated lazy slide catalog. Its traversal
iterates the retained slide metadata directly rather than collecting slide
handles or slide text. For each logical slide it loads one selected `PartData`,
parses the bounded semantic text, writes it, and releases the traversal's
current value before continuing.

The source cache may retain previously loaded payloads. That retention is
governed by the caller's configured cache limits and is not eliminated by the
sink API. Consequently this change claims one bounded current-slide
raw/processed/text state, not absence of all previously loaded slide payloads
or a whole-transaction memory bound.

Source execution and source-version checks surround opening, selected-slide
loading, parsing, sink acceptance, and finalization. If cancellation or a
source revision is observed after output has been accepted, the operation
returns a document error with the exact partial progress; that source error
takes precedence over a competing sink result. Accepted sink bytes are not
rolled back.

## Declared-size and parser bounds

Source-backed part loading treats the declared size exposed by an untrusted
`PartView` as attacker-controlled metadata. A declared size above the
effective read/decompression limit is refused before decompression or payload
materialization. After the read, the actual payload size is checked as well,
so an inaccurate declaration cannot bypass the same bound. These checks are
resource guards, not evidence of physical I/O or decompressor work avoided.

Semantic slide parsing retains the existing finite raw XML, processed XML,
decoded text, markup-compatibility, and XML-depth limits. A limit or parse
failure is returned as a typed document error carrying the exact progress
already accepted by the sink.

## Focused evidence

Focused sequential-text tests cover eager/source text parity, configured
empty-slide and separator behavior, output/object limits, short/zero/
over-reporting/interrupted/failing sinks, malformed later slides, source
staleness, cancellation precedence, declared parser ceilings, and the
pre-decompression declared-size refusal with post-read verification.

No new perf-baseline selector is added. Existing full-text PPTX timing remains
the aggregate-string control and must not be presented as sink evidence. This
record reports deterministic API/resource behavior only; no latency,
throughput, allocator, RSS, physical-I/O, decompression, zero-copy, or
whole-transaction memory claim follows.

## Scope boundary

`performance_claim: none` is authoritative for this change. Sink byte/object
reports and source-cache/read counters, when inspected in a focused replay,
are logical deterministic observations. They do not establish physical I/O,
allocation, RSS, decompression, or comparative speed. A future quantitative
claim would require a separate selectable case and reproducible comparison
with explicit cache policy and independent semantic output oracle.
