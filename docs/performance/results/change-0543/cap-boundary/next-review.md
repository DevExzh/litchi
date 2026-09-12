# 0543 cap-boundary next proposal review

This is an independent static review of the unmeasured follow-up described in
[`next-design.md`](next-design.md). The prepared source fragment is the
`next-candidate-src` tree and is based on the frozen 0543 candidate. The
measured 0543 source and its evidence remain immutable; this review changes no
Rust source, build, test, or measurement artifact. A follow-up patch has not
been generated at review time.

## Static disposition

The delimiter bound is sound as a conservative upper bound under the pinned
`quick-xml 0.41.0` reader call and its current configuration. I found no
concrete undercount for any current `Event` variant, BOM handling, text
trimming, declarations, comments, CDATA, references, or malformed tails. The
candidate may generate a provenance-bound follow-up patch, conditionally on
the proof and test obligations below. It is not approved for measurement or
runtime retention until those obligations and the fresh performance gates
pass.

## Control-flow effect

`shared_event_bound_within_cap` runs in
`raw::worksheet::parse_source_with_observer` before it constructs either
`NsReader` or the raw worksheet parser. A `false` result returns
`ProvisionalFailed`; the caller drops its empty validator and executes the
established full validator followed by the authoritative raw parser. No
validation is skipped, no provisional store is published, and source,
MCE/x14ac, error-order, and publication fences remain unchanged.

If the bound fits, the existing one-reader path and its runtime event-cap
defense are unchanged. The runtime cap remains necessary because this lexical
preflight is only an upper-bound proof and the reader/parser configuration is
an implementation dependency. The added `large_enum_variant` expectation is a
narrow lint disposition for the private owned `Complete(Result<Store>)` arm;
it does not alter ownership or callback lifetimes and still requires a fresh
quality receipt.

The proposal removes the current valid-input cliff for sources whose lexical
bound exceeds 131,072: those sources bypass the discarded shared prefix and
pay the normal two authoritative traversals. Sources whose bound fits retain
the shared one-traversal success path. The two full `memchr2` scans add work to
every admitted source, so the 96 × 96 and 128 × 128 success controls must be
measured; the supplemental 160 × 160, 164 × 164, and 256 × 256 cases must
confirm that early refusal is no slower than the restored path within the
separate cap-boundary gate.

## Event-bound proof

The helper starts with one count for the terminal `Event::Eof`, counts each
`<` and `&`, and counts a possible text start at the source beginning and
after each `>` or `;` followed by a byte other than `<` or `&`. For the pinned
`NsReader::from_reader(content)` followed by `read_event()`, the needed
implication is:

```text
actual emitted events > MAX_SHARED_PROVISIONAL_EVENTS
    => lexical bound B(content) > MAX_SHARED_PROVISIONAL_EVENTS
```

Every markup event (`Start`, `End`, `Empty`, `Comment`, `CData`, `Decl`, `PI`,
or `DocType`) is initiated by a counted `<`. With
`expand_empty_elements = false`, a self-closing tag emits one `Empty` event;
there is no synthetic `End` event requiring a second count. Every successful
`GeneralRef` is initiated by a counted `&`. A dangling or malformed reference
under `allow_dangling_amp = false` returns a reader error instead of an
uncounted event, so counting its `&` is conservative.

The only remaining event kind is `Text`. With the current
`trim_text_start = false` and `trim_text_end = false`, an initial text segment
is covered by the initial-byte term. A text segment following markup begins
after the markup's terminating `>`, and one following a successful reference
begins after its `;`; the second scan counts the terminator whenever a
non-`<`, non-`&` byte follows. If the next byte is `<` or `&`, the reader emits
the next markup/reference event or reports its malformed-reference error
without an intervening text event. If the delimiter is at EOF, the reader
returns `Eof` rather than an empty text event under the current defaults.

`NsReader::resolve_event` only annotates existing `Start`, `Empty`, and `End`
events; it does not create events. The reader removes a leading UTF-8 BOM
during initialization. Since the first BOM byte is neither `<` nor `&`, the
initial term counts one extra possible text event, which is safe. XML
declarations and processing instructions are markup events and are covered by
their `<`; declaration encoding does not add events in the pinned default
dependency configuration.

Delimiters inside attributes, comments, CDATA, doctype bodies, or ordinary
text can inflate the bound. A `>` or `;` that is not an actual event terminator
can only add a false positive, and a `<` or `&` embedded in one of those
regions is likewise over-counted. This can route a source to authoritative
fallback but cannot admit an over-cap event stream.

The proof depends on keeping `allow_dangling_amp = false`,
`expand_empty_elements = false`, and the current `read_event()` API. Turning
on empty-element expansion or dangling-amp recovery would add event forms
that this bound does not cover. Any change to reader trimming, encoding
features, dependency version, or the `NsReader` call must trigger a new proof
and edge review. `check_end_names = true`, BOM removal, and namespace
resolution do not weaken the bound.

## Edge coverage required before measurement

The prepared private tests prove the bound admits the 96 × 96 and 128 × 128
numeric grids and declines the repeated-comment cap fixture. They do not yet
serve as a complete proof oracle. Before generating or measuring a retained
follow-up, add a focused bound test that compares `B(content)` with direct
`quick_xml` event counts for:

- empty, whitespace-only, and BOM-prefixed sources;
- XML declarations and processing instructions;
- comments, CDATA, and doctype declarations, including internal `>` and `&`;
- quoted `>` and `&` in attributes;
- ordinary text containing `>` and `;`;
- chained text/reference segments, named and numeric references, and dangling
  or malformed references;
- nested, self-closing, and malformed-tail markup;
- exact bound `131,072` and first over-bound `131,073` outcomes.

The oracle must use the same `NsReader` configuration as the shared path and
count EOF. It must establish both that every source with actual events above
the cap returns `false` and that an exact bound returns `true`. False
positives are expected and should be asserted as safe fallback rather than
treated as proof failures. The source-eligible cap fixture should also retain
its public error, retry, source-identity, and no-op publication checks.

## Remaining admission conditions

The next patch should remain limited to the prepared five-file inventory and
must bind its complete source manifest and patch hash. It must preserve the
existing source-byte aggregate budget, 8 MiB eligibility fence, UTF-8 and
MCE/x14ac marker checks, validator-first fallback, post-EOF raw result
forwarding, x14ac retry, higher-ranked event callback, and owned source
payload. The helper must remain before provisional reader/parser construction
and must not be reused with another reader configuration.

Fresh baseline/follow-up captures are required for the real 96/128 primary
shapes, all existing valid/late-validator/late-raw guards, the supplemental
160/164/256 boundary sizes, and the ordinary eager-read controls. The byte
preflight's scan cost, early-decline cost, allocator behavior, and repeat
stability need separate dispositions. No 0542 or 0543 timing or allocation
result is evidence for this unmeasured proposal. A candidate that fails the
edge oracle, changes any error/source/publication result, or exceeds the
unchanged valid/planning/allocation gates must be rejected and restored.

This proposal remains within the OLE2/OOXML priority. ODF optimization stays
deferred until that goal is complete.
