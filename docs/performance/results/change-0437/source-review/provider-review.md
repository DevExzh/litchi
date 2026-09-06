# 0437 ODP plain-slide provider review

This is a source-only review of the provider draft in
`/tmp/litchi-goal-0437-odp-provider/streaming.rs` and `api.md`, checked against
the integrated `crates/litchi-odp/src/streaming.rs` and the existing ODP
Builder/XML/common-writer code. I did not run a build, test, script, formatter,
or benchmark. The integrated file differs from the draft only by the reported
cleanup/formatting and the removed unused conversion helper.

## Review result

I found no production-blocking mismatch in the fixed fresh-slide grammar. The
provider emits the Builder no-transition prelude, the default `pageN`/`dp1`/
`Default` page attributes, title and body frames in the right order and with
the right geometry, and the same paragraph/control/entity spelling as
`authoring/builder/xml.rs`. The important edge semantics are represented:

* `Some("")` still emits one title frame and one empty P1 paragraph;
  `None` omits that frame.
* An empty body omits the body frame. A body-only page uses `2.0cm`; a page
  with a present title uses `5.0cm`.
* LF splits paragraphs, CR emits `text:line-break`, tabs emit `text:tab`, and
  leading/trailing/repeated spaces use `text:s` with the Builder's literal
  single-interior-space rule.
* XML 1.0-invalid scalars are refused before the slide is admitted, while
  ordinary XML-significant characters use the same five escaped entities as
  the Builder.

The source is consumed incrementally and only one bounded page fragment is
retained. The source, slide, content-XML, archive, output, memory, and common
XML-audit ceilings are checked before the corresponding admission. The sink
adapter reports the accepted caller prefix, retries short writes through
`write_all`, preserves source and sink causes, and leaves failed publication
outputs explicitly disposable. The provider's modeled memory reservation is
consistent with its stated scope rather than claiming ZIP/compressor/auditor
peak.

## Claim-alignment findings

### Work and resource dimensions need a precise public scope

The code consumes `Resource::Work` for the fixed content shell, every emitted
authored content fragment byte, and the fixed styles/meta members. It does not
consume Work for the MIME member, ZIP local/central headers, the generated
manifest, deflate/compression effort, or common XML-audit/parser work. That is
an acceptable bounded provider policy, but `api.md` should state it explicitly
if reports expose Work evidence; a reader must not interpret it as complete
CPU or ZIP-publication work.

The other dimensions have similarly specific meanings and should remain
explicit in the harness schema: `InputBytes` is raw UTF-8 title/body bytes,
`Objects` is one admitted source slide, `OutputBytes` is the accepted caller
sink count, and `Memory` is the fixed retained-provider model. XML depth is
enforced by `XmlAuditLimits`; the execution-context `Resource::Depth` budget
is not charged by this sequential provider. Cancellation is cooperative at
source/item, fragment-write, and scalar/span preparation boundaries; a
concurrent cancellation can allow one already-admitted bounded ordinary-text
span to finish.

### Raw text counters and lexical XML counters are different domains

The provider's title/body/aggregate limits and report counters use raw source
UTF-8 bytes. The common XML audit counts the composed lexical document, where
`&`, `<`, quotes, and apostrophes expand to entity bytes and the fixed shell
also contributes. `content_xml_bytes` is the composed content member including
the shell. These values must not be compared as if they were one input-byte
projection; a source can satisfy the raw text limit and still hit the lexical
XML text or byte limit.

### Invalid input currently preserves the category but flattens its value

`ProducerState::invalid_error` stores `error.to_string()` and later rebuilds an
`Error::InvalidFormat`. All current invalid producer paths use that variant,
so this does not lose a category today. If the public promise to preserve
typed failures is intended to cover future invalid variants or structured
fields, retain the `Error` itself alongside the producer/source and return it
directly. Fallible source errors and publication/sink errors already retain
their typed chains and accepted-byte field.

### The slide-limit lookahead should be documented

After `max_slides` have been emitted, the callback polls one additional source
item to distinguish exact exhaustion from an over-limit source. This is bounded
to one lookahead and is useful for reporting an actual overflow, but
“consumed once” in the API should mention that the source may be polled once
more when checking the slide ceiling. A source with side effects should not
assume that the extra item is never requested.

### Minimum output ceiling is only a ZIP lower bound

`StreamingLimits::new` accepts `max_output_bytes >= 22`, which is enough for an
empty ZIP end record but not enough to publish even the required ODP members.
The resulting runtime `LimitExceeded` is coherent, and retaining the small
constructor lower bound is useful for limit tests, but the API should say that
the constructor does not guarantee successful package publication for every
finite output ceiling.

## Test and oracle recommendations

The focused coverage should keep the existing byte-for-byte Builder control
for a small corpus, plus an independent XML/package oracle for formal shapes.
The edge corpus should include empty title, absent title, empty body, leading
and trailing spaces, repeated spaces immediately around tabs/CR, mixed CR/LF,
all five escaped characters, non-ASCII text, and an invalid XML scalar. It
should assert exact report counters and the page/frame attribute order, not
only a reopened text projection. The package oracle should pin the five-member
set, mimetype, compression/topology, manifest bindings, and the documented
styles/meta byte hashes. It should also keep one-parent/one-child exact and
one-under checks for Memory, InputBytes, Objects, Work, and OutputBytes, and
verify source/sink error-chain markers plus accepted partial output.

Do not import the ODT provider's adjacent-control lexical variation into ODP
without changing the stated contract: the ODP API promises Builder-equivalent
bytes, and the current encoder matches that promise. Add the control-adjacency
cases to prevent an accidental future “normalization” from changing the
cross-role lexical result.

## Disposition

The integrated provider is suitable for the planned focused validation and
formal oracle, subject to the scope prose above. No source edit was made by
this review.
