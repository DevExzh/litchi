# 0435: bounded ODT paragraph authoring

The previous batch made verified progress: production `65dfb1b1` and evidence
`76710f5c1` were committed, copied-bundle verification passed, and its task
temporaries were removed. The new baseline starts from
`76710f5c1676fd305e85e6d045dfeb570eb1f421`. The accepted ADR tree remains
`c950b6c8be822561b498d7bbe87c460873dcbf49`; `docs/GOAL.md` is unchanged.

## Measured-path preparation

The existing ODT builder retains its paragraph model, constructs a complete
body and `content.xml`, and returns a complete archive vector. The old
`OdtSemanticCreateSmall` case always creates its tiny shape, so it cannot
establish resource growth for large fresh authoring. Add a separate opt-in
buffered selector over 64/8,192/32,768 paragraphs, capture its baseline before
applying the production candidate, and keep the old selector intact.

The new writer should consume ordered paragraph inputs once and publish one
complete bounded paragraph fragment at a time through the existing common
generated-XML method. It should retain the builder's five-member default ODT
topology and fixed default metadata/styles. Each role must emit deterministic
archive bytes. Streaming ZIP descriptors may differ from buffered framing;
cross-role correctness therefore requires equal supported paragraph semantics
and explicitly checked XML/topology identities, without inventing raw archive
equality. Timers must include paragraph generation and publication to the
hashing discard sink. Corpus construction, reopen/oracle checks and process
probes remain outside the operation timer.

The common generated-XML envelope accepts nested opening tags in its prefix,
not completed sibling elements. The streaming content shell therefore omits
the Builder's three optional empty scripts/font-face/automatic-style children.
Both roles use UTF-8, but that encoding label does not imply lexical equality.
Cross-role content XML hashes may differ; semantic paragraphs and the complete
fixed styles/meta member identities must agree. No common envelope expansion
or full-content buffer is needed for this fresh plain-text grammar.

## Whitespace is part of the semantic contract

[ODF 1.3 Part 3](https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part3-schema/OpenDocument-v1.3-os-part3-schema.html),
sections 6.1.2–6.1.5, applies whitespace processing after XML decoding. Literal
tab, CR and LF become spaces; edge spaces disappear and runs collapse. ODF
space, tab and line-break elements are interpreted after those reductions.
The new writer must encode spaces and LF/tab controls so expected text survives
that process. Numeric CR references alone do not establish native CR fidelity.
Distinct CR input needs typed refusal or an explicit caller-selected policy;
silent normalization is not an acceptable default.

The buffered comparison corpus uses Unicode, XML metacharacters and single
interior spaces, avoiding the current builder's raw-whitespace ambiguity.
Provider correctness tests separately cover leading/trailing/repeated spaces,
whitespace-only and empty paragraphs, LF, tabs, CR refusal, UTF-8 boundaries,
and invalid XML controls. Zero paragraphs and one empty paragraph are distinct.
The fixed grammar admits no caller XML or attributes. Common lexical auditing
does not substitute for the provider's ODT grammar responsibility.

## Resource and failure boundaries

Align finite limits with the ordinary reader's one-million text blocks,
64 MiB aggregate text and 256 MiB content ceilings. Use explicit per-paragraph
text and encoded-fragment windows, aggregate text/XML/output limits, checked
arithmetic, hierarchical Memory/Work/OutputBytes budgets, per-scalar
cancellation polling and typed accepted-output progress. The reusable fragment
window is not a total allocator or RSS bound.

The common writer admits ZIP framing before pulling the first paragraph.
Producer, validation and sink failures may therefore leave an incomplete
caller sink. Preserve the original typed cause, report only acknowledged sink
bytes, discard the failed package and never finalize it after a refusal.
Include a fallible paragraph-source surface rather than silently truncating an
upstream producer failure.

Tests must cover exact/one-under limits, parent/child budgets and memory release,
ordered/fallible sources, cancellation at several publication phases, short and
partially failing sinks, aggregate audits and deterministic semantic reopen.
Root serializes all build/test/profile jobs and owns staging and commits.

The broader non-iWork goal remains active. Rich authoring, existing-document
append, package-Part addition, arbitrary repackaging, native producer breadth,
cold/range I/O and explicit scaling remain separate required work.

## Buffered harness gate

Commit `7c25f299d` adds the opt-in scalable buffered case. Before production
changes, the ODT release suite passed 982 tests and the standalone release
harness passed 313 tests with one existing ignored test. The harness oracle
checks the five members, stored mimetype/Deflate XML methods, four manifest
bindings, exact paragraph and full-text projections, and independently pinned
complete default styles/meta hashes. It also binds the corpus manifest and
retained target payload to reopened archive identity before timing.

Strict harness Clippy still fails on 29 pre-existing diagnostics. The retained
comparison proves an identical message/file multiset to batch 0433, no finding
on changed `lib.rs` lines, and none in the new helper. This is a scoped gate,
not an unqualified strict-Clippy pass. Format and crate-boundary checks pass.
Failed initial compile attempts remain in `checks/` with terminal receipts.

The six preparatory pilots and both large whole-process profiles passed before
production edits. Across 64/8,192/32,768 paragraphs, timed allocation calls were
1,053/112,819/450,741 and region peak above entry was
458,819/5,616,425/22,450,985 bytes, identical across each shape's three allocator
samples. All operations released their live allocations before the endpoint.
These observations support testing bounded publication; they are not formal
latency, causal allocation attribution, RSS, or speedup claims. Raw inputs and
scope limitations are bound by `buffered-hypothesis.json`.

## Typed fixed-XML publication seam

The existing common `add_file` compatibility method converts transport errors
to strings. Its opaque reader method refuses XML, while generated fragments
refuse comments. Builder's exact default styles contains a comment, so neither
alternative preserves both its bytes and a nested sink error. Add one common
typed authored-XML slice publisher that reuses existing XML/path/manifest
validation and the sized Deflate writer, returning `PackageWriterError` without
stringification. This does not change the accepted XML grammar or the legacy
method. The family provider uses it only for its fixed styles/meta members.
Tests must prove exact comment bytes, fail-closed validation, transport limits,
poisoned finalization, and nested caller error preservation.

## XML limit attribution

Final source review found that the existing generated-member reader converted
lexical and aggregate XML audit exhaustion to strings. The common seam now
retains a typed `GeneratedXmlLimitExceeded` through its existing core and I/O
error boundaries. `PackageWriterError::xml_limit()` recovers that attribution
without diagnostic parsing; its source traversal is bounded and includes owned
I/O errors. The ODT provider reports the XML resource, first observed value,
configured ceiling, and its sink adapter's acknowledged output count. Malformed
XML retains its existing refusal mapping. The change adds no successful-path
validation pass and does not widen the exhaustive archive-resource enum.

Paragraph, content, and XML-audit ceilings are independent. In particular,
`max_paragraphs` is an admission ceiling, not a promise that that many paragraphs
fit every event/text/output limit. The modeled provider Memory reservation is
also separate from measured total allocator peaks and whole-process RSS.
