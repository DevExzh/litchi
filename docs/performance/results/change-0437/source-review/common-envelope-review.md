# 0437 common generated-XML seam review

This is a review of the opt-in common seam described in
`/tmp/litchi-goal-0437-odp-audit/design.md`. It is a test/contract plan only;
it does not authorize changing the production tree. The existing
`GeneratedXmlEnvelope::try_new(prefix, suffix)` contract must remain exactly as
it is: its prefix contains only an optional declaration and open elements, its
suffix contains only matching end elements, and an end element in the prefix
continues to be rejected (`generated_xml.rs:76-114, 562-654`).

## Proposed opt-in envelope grammar

The new constructor can have any narrow name, but its inputs should have these
meanings:

```text
prefix = declaration?
         root/open-ancestor events
         balanced fixed child subtrees*
         final insertion-path Start events+
suffix = matching End events for the still-open stack, in reverse order
output  = prefix + nonempty fragments* + suffix
```

The root and every element in the final insertion path remain open at the
prefix boundary. A balanced child is a complete `Start ... End` subtree or an
`Empty` element that is completely before the final insertion path. Balanced
children may nest and may have their own attributes. Once the final insertion
path starts, the prefix may not close an element or open a second fixed child;
this prevents the API from becoming an arbitrary XML concatenator. At least
one final `Start` must remain open, so a fixed complete document cannot be
passed as an insertion envelope.

The common layer is lexical, as its module documentation states
(`generated_xml.rs:1-10`). “Namespace-aware lexical matching” therefore means
matching the complete raw qualified name (`prefix:local`, or the unprefixed
name) byte-for-byte between every fixed `Start` and `End`. It does not resolve
namespace URIs or decide whether an ODP QName is legal. A different prefix for
the same local name is still a mismatch at this layer; the provider owns URI,
declaration, and schema validation. Prefix/default namespace bindings must be
declared before use and remain the provider's responsibility.

The parser must enforce the following for both the fixed balanced portion and
the final path:

* There is one document root. A balanced child cannot occur before a root,
  after the root has closed, or as a second top-level root.
* An optional XML declaration is allowed only once, at byte zero. A second
  declaration, a DTD, processing instruction, or comment is rejected. The
  suffix has only `End` events; it has no declaration, text, whitespace,
  `Start`, or `Empty` event.
* The prefix/suffix byte boundary is between parser events. A boundary inside a
  name, attribute, quoted value, `>`, `/>`, UTF-8 sequence, or any other event
  is rejected before a producer can run.
* Every suffix end tag matches the remaining open stack in reverse order. An
  extra end tag, missing end tag, local-name mismatch, or qualified-name
  prefix mismatch is rejected. The stack must contain exactly the root/path
  elements at the boundary and be empty after the suffix.
* `xml:space` is rejected on every fixed `Start`, including a nested balanced
  child and every final-path start. There must be no inherited whitespace mode
  crossing into a generated fragment. This preserves the current
  `reject_xml_space` rule (`generated_xml.rs:656-667`) rather than applying it
  only to the root.
* Character data is only allowed where it is fixed content inside a balanced
  child. Text, CDATA, or a predefined/numeric reference between top-level
  events, before the root, after the root, or in the final insertion path is
  “unexpected text” and is rejected. If the implementation chooses the even
  narrower ODP policy of rejecting all fixed character data, that must be an
  explicit rule rather than an accidental parser behavior. For the broader
  lexical rule, admitted fixed text/reference bytes are included in the shell
  report; undeclared named references remain rejected. In either policy,
  indentation/text around the dynamic insertion point must not be silently
  inherited.
* Fragments retain the existing one-complete-element-root contract in
  `validate_fragment_shape` (`generated_xml.rs:670-748`): no declaration,
  DTD, comment, PI, outside-root text/reference, or multiple roots. The
  provider, not this seam, verifies that the root is `draw:page` and that its
  children are supported.

The fixed shell should be validated as one complete XML document for lexical
well-formedness after the balanced prefix and suffix have been interpreted.
The implementation may retain bounded copies of the fixed prefix/suffix, but
must not materialize the composed member or retain all fragments. In
particular, moving `automatic-styles` into only the first callback result would
make later page fragments sibling roots and violates the design contract.

## Report and limit invariants

The shell report is the audit of the complete fixed prefix plus suffix,
including every balanced child. It must include:

* all fixed bytes and attributes;
* all fixed parser events, including one document EOF event and any `Empty`
  events;
* all admitted fixed character-data bytes; and
* the greatest depth anywhere in a balanced child or the open path.

`GeneratedXmlReader::new` currently seeds the composed report from the shell
and checks the requested limits before allocating the fragment buffer
(`generated_xml.rs:205-238`). The opt-in path must keep that order and must not
seed from only the old open-path shell. For every nonempty fragment, preserve
the current arithmetic (`generated_xml.rs:296-339`):

```text
events       = shell/events + Σ(fragment.events - 1)  // one shared EOF
bytes        = shell/bytes + Σ(fragment.bytes)
attributes   = shell/attributes + Σ(fragment.attributes)
text_bytes   = shell/text_bytes + Σ(fragment.text_bytes)
max_depth    = max(shell/max_depth,
                   insertion_depth + fragment.max_depth)
fragments    = number of accepted nonempty producer results
```

All additions, including insertion depth plus fragment depth and fragment
count, are checked. A balanced child deeper than the insertion path must still
win `max_depth`; a fragment nested below the path must add the path depth. No
fixed bytes or attributes may be counted twice when the first fragment is
read. An empty producer result is EOF and contributes no fragment counters.

The inclusive caller limits and immutable hard ceilings remain the same
resources (`Bytes`, `Depth`, `Events`, `Attributes`, `TextBytes`, plus the
per-fragment lexical/token limits). Check the complete fixed shell against the
caller limits before the first producer callback. Check each fragment's own
audit before adding it to the composed report, then check the checked composed
report before exposing any bytes from that fragment. A one-under limit must
report the exact first exceeding value through `GeneratedXmlLimitExceeded`,
not a generic ZIP or producer error. The existing public limit tests establish
this inclusive/one-under convention (`tests/generated_xml_limits.rs:139-284`).

The shell and one reusable fragment window are bounded allocations. There must
be no full composed XML allocation in the new path; output is still streamed
as fixed prefix, each accepted fragment, then fixed suffix. This is an
implementation review invariant even if output-equality tests cannot prove the
allocation shape.

## Error and publication order

The observable order should be stable and preserve the existing reader/writer
boundary:

1. The opt-in constructor checks checked lengths/hard ceilings, UTF-8 and XML
   event syntax, boundary alignment, balanced-prefix grammar, exact end-name
   matching, `xml:space`, and the complete fixed-shell audit. It returns before
   ownership of a producer or any output publication on failure. The old
   `try_new` retains its current audit/shape behavior and error compatibility.
2. `prepare` validates caller limits and nonzero fragment capacity, audits the
   fixed shell under those limits, and reserves the reusable fragment buffer
   before invoking the callback (`generated_xml.rs:205-238, 817-867`). A shell
   limit failure therefore has zero producer calls.
3. `PackageWriter::add_generated_xml` validates MIME/path/XML classification,
   manifest collision, encryption/signing state, and archive admission before
   the first callback (`writer.rs:1246-1293`). The existing test
   `rejected_archive_budget_does_not_pull_the_first_fragment` is the required
   ordering oracle (`writer_generated_xml_tests.rs:124-151`).
4. Once the content local header has been admitted, the first callback may
   run. A producer error is retained as the source error, reports the accepted
   sink prefix, poisons the package, and makes finish fail. The same holds for
   a later producer error after one or more accepted fragments; no generic XML
   error may replace the source cause. Existing first/later failure tests cover
   this (`writer_generated_xml_tests.rs:72-122`).
5. Within `fill_fragment`, preserve the current precedence
   (`generated_xml.rs:260-294`): propagate a producer `Err` first; if the
   callback claimed success but the bounded scratch writer failed, return that
   buffer failure; reject `false` with bytes and `true` with no bytes; audit and
   shape-check a nonempty fragment; only then add checked aggregate counters
   and enforce aggregate limits. A failed fragment is not emitted and no later
   callback is pulled.
6. The suffix is emitted only after an empty `false` result. Any failed reader
   remains failed; it cannot later emit a shell or suffix. Sink short writes or
   transport failures preserve the exact accepted count and poison finalization,
   as in `short_sink_failure_reports_accepted_progress`.

This ordering intentionally distinguishes constructor/prepare failures (no
content entry and no source call) from callback failures (a content entry and
possibly a nonzero accepted prefix). It also keeps lexical fragment failures
ahead of aggregate-limit failures when both could apply, because the fragment
must first be independently audited.

## Concrete common-seam tests

Add focused tests alongside `generated_xml_tests.rs`; exercise the writer
boundary in `writer_generated_xml_tests.rs` and, where a typed limit is needed,
reuse the helper pattern in `tests/generated_xml_limits.rs`.

### Shape and compatibility

1. **ODP balanced prelude acceptance.** Use a prefix equivalent to:
   `<?xml ...?><root><scripts/><styles><style id="dp1"><leaf/></style></styles><body><presentation>`
   and suffix `</presentation></body></root>`. Emit two complete fragments such
   as `<page id="1"/>` and `<page id="2"/>`; assert exact byte order, one root,
   and no prelude duplication. Also test zero fragments and assert the fixed
   shell is emitted exactly.
2. **Nested fixed accounting.** Put a text-bearing balanced child and an
   attribute-bearing child before the path. Assert the report equals a direct
   audit of the assembled expected bytes, including text, attributes, events,
   and the deeper fixed max depth. Add a fragment nested two levels below the
   insertion path and assert path-plus-fragment depth wins when appropriate.
3. **`Empty` is opt-in only.** A balanced `<scripts/>` is accepted by the new
   constructor. The same prefix passed to old `try_new` remains rejected,
   proving that the strict API was not widened.
4. **QName matching.** Reject `a:root ... </b:root>`, local-name changes,
   reversed nested ends, extra ends, and a suffix that leaves an ancestor open.
   Do not resolve equal namespace URIs to make different lexical prefixes pass.
5. **Boundary and root checks.** For the same valid fixture, put the boundary
   inside a start tag/name, quoted attribute, `/>`, UTF-8 code point, and final
   start tag; reject each. Reject a balanced element before the root, a closed
   root followed by another root, and a prefix with no final open insertion
   path.
6. **Forbidden events/state.** Reject DTD, comment, PI, second declaration,
   suffix start/empty/text events, `xml:space` on a nested fixed child and on
   the final path, and text/reference outside a balanced element or in the
   insertion path. If fixed text is admitted by the selected grammar, include
   one positive in-element text case and one negative sibling/outer text case.
7. **Existing strict cases stay strict.** Keep the current assertions for
   mismatched suffixes, inherited `xml:space`, and prefix text
   (`generated_xml_tests.rs:237-242`) against `try_new`; add equivalent negative
   assertions against the opt-in constructor without changing the old test's
   meaning.

### Counters and limits

8. **Direct report oracle.** For the accepted ODP fixture, concatenate the
   expected fixed prefix, fragments, and suffix only in the test, run
   `xml_minifier::audit::verify_authored`, and compare every public report
   field. This catches omission of balanced child bytes/attributes/text and
   double-counting of EOF or fixed shell data.
9. **Inclusive/one-under shell limits.** For each of `Bytes`, `Depth`,
   `Events`, `Attributes`, and `TextBytes`, set the limit to the exact report
   value and to one below it. Include a fixed child whose depth exceeds the
   final path and an attribute/text-bearing fixed child. The exact case must
   publish; the one-under case must return the typed resource/actual/maximum
   failure before any callback.
10. **Aggregate fragment limits.** Use two individually valid fragments and set
    each composed limit to the first-fragment total and one below the
    two-fragment total. Assert the second callback is entered, its fragment is
    not emitted, the report remains at the first accepted total, and the error
    has the exact aggregate resource/value. Repeat for bytes, attributes, text,
    events (subtracting one EOF per fragment), and depth. Keep the existing
    four aggregate tests as regression oracles.
11. **Per-fragment capacity/token limits.** Make a fragment independently
    exceed its reusable byte/token window while the fixed shell and aggregate
    limits allow it. Assert the scratch failure is returned before output from
    that fragment and later callbacks are not called.

### Publication/error ordering

12. **No callback on preflight refusal.** Through `PackageWriter`, pass a valid
    balanced envelope but exhaust entry/archive admission; assert zero callback
    calls and only the previously accepted MIME bytes. Pass malformed balanced
    shape and an over-limit fixed shell similarly; assert no content local
    header or producer call.
13. **Source failures.** Make the first callback fail and then make the second
    callback fail after one valid fragment. Assert the original error remains in
    the chain, accepted progress is nonzero, no suffix is treated as complete,
    and `finish_to_writer` is refused.
14. **Producer contract.** Exercise `true` with an empty fragment and `false`
    after writing bytes. Both fail before any such fragment is emitted. A clean
    `false` emits the suffix exactly once; a subsequent read returns the
    latched failure/EOF according to the existing reader state contract.
15. **Transport progress.** Use the existing short sink and assert the accepted
    count covers exactly the sink's accepted prefix even when the prefix,
    balanced child, fragment, or suffix crosses a write boundary. Finalization
    remains refused after the transport error.

The acceptance tests should verify only the format-neutral seam. ODP-specific
tests must additionally assert namespace bindings and `draw:page`/child
semantics; the common constructor must not grow those policy checks.
