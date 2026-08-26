# Change 0294: DOCX semantic text sink

## Scope

The DOCX owner now exposes bounded sequential semantic text output on the
eager `Document` facade and on `source_backed::Package`. Output is one
`WordprocessingML` paragraph object at a time, in XML order, including
paragraphs in table cells. `TextOutputOptions` controls paragraph separators,
empty paragraphs, output bytes, and object count. `TextOutputReport` reports
exact bytes accepted and complete objects emitted; parser, limit, and sink
errors preserve that partial progress.

The parser has one current paragraph working string. Its raw and processed XML,
event, attribute, namespace, depth, paragraph, run, reference, per-event
decoded-text, current-paragraph, and aggregate decoded-document ceilings are
independent bounded resources. Every string growth is fallible. The parser
recognizes canonical transitional and Strict WordprocessingML names, preserves
the established tab, line-break, no-break-hyphen, soft-hyphen, CDATA, and
predefined/numeric-reference text semantics, and rejects malformed XML,
foreign text lookalikes, unsafe namespace state, DTDs, processing
instructions, and text outside the document root.

Eager documents parse their already-visible MCE view. Source-backed packages
check execution and source freshness before deferred metadata/payload work,
preflight the central-directory declaration before reading the main payload,
retain the payload reader's actual-length and TOCTOU checks, and use the
existing managed-source MCE refusal. A source-checked sink fences every
underlying write and source/cancellation errors take precedence over a
simultaneous parser, limit, or sink error.

## Performance claim

`performance_claim: none`.

This change claims only a bounded semantic working set: one current paragraph
text value plus parser and namespace state. It does not claim constant total
RSS or a particular latency, throughput, allocator, physical-I/O,
decompression, or zero-copy result. The retained eager/package payload and
the source payload reader remain outside the one-paragraph working-set scope.
For eager callers, `DocumentPart` construction and visible MCE preprocessing
also occur before this sink and are outside that scope.

## Exclusions

The sink does not retain or interpret headers, footers, media, `altChunk`
payloads, rendering/layout state, field evaluation, or unsupported package
parts. It does not alter legacy `text()`/`extract_text()` behavior, flush or
roll back the caller-owned sink, or provide input-streaming, allocator, or
physical decompression guarantees.
