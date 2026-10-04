# 0905 XML-audit attribute-summary source review

This review covers the current 0905 production candidate in
`crates/xml-minifier/src/audit.rs`. The candidate is compared with the frozen
baseline copy at `docs/performance/results/change-0905/baseline-audit.rs`.
The baseline copy has SHA-256
`91424b57dfcecaf6cb204429e6d779a0c81a8b55c6546d616c694f4ea2455ac0`; the
candidate file reviewed here has SHA-256
`eddc437156c976c55ce36b0bf8f0ef47cc4943d296735a1ac8d6ef120b8754e7`.
The repository was at `4f39f5a88d` while this review was written. No build,
test, workload, profiler, or production edit was performed by this review.

## Conclusion

The candidate has a defensible semantic boundary for a differential gate. I
found no source-level counterexample that bypasses a required XML check or
changes a successful report within the boundary it implements:

* zero attributes skip the second iterator and preserve the inherited
  `Space` and aggregate counter;
* one raw attribute whose key is not exactly `xml:space` skips the iterator and
  performs exactly one aggregate-attribute charge at the existing tag offset;
* `xml:space` and every tag classified as `Many` retain the existing checked
  iterator, duplicate detection, second per-element bound, and value
  normalization path.

This is a source qualification only. It does not authorize production
adoption or establish a speedup. Exact differential tests, quality checks,
and the fresh paired native gate remain required.

## Why the fast boundary is sound

`scan` and the streaming reader both call `check_start` before
`inspect_attributes` for `Start` and `Empty` events. `check_start` now returns
`AttributeSummary`, but it still completes `check_attribute_layout` first.
That pass reads every separator, name, equals sign, quote, and value. Under
`Policy::SOURCE` it also completes the existing namespace declaration and
expanded-name checks before returning. The callback only records the raw key
classification; it does not make an early return.

Therefore a second attribute, including a malformed second attribute, cannot
be classified as `One` and reach the fast return. A malformed tail after the
first attribute similarly causes `check_attribute_layout` to return before
`inspect_attributes` sees the summary. This preserves the prior error order
for duplicate, malformed-tail, missing-separator, and invalid-value cases.

The key test is deliberately the same raw-byte comparison as the old
`attribute.key.as_ref() == b"xml:space"`. It does not use namespace expansion,
case folding, or decoded values. A namespace declaration, prefixed attribute,
Unicode or punctuation key, and `xml:lang` therefore follow the one-attribute
fast case when the complete first pass accepts them; only the exact raw
`xml:space` spelling takes the value-decoding path.

For a zero-attribute tag, the old checked iterator yields no item and changes
no state. Returning `inherited` is equivalent. For one accepted non-
`xml:space` attribute, the old iterator yields one item, performs no duplicate
comparison that can fail, and the body only increments the aggregate
attribute counter. The candidate performs that same `checked_add` with the
same `Resource::Attributes`, limit, amount, and tag-start offset. It does not
remove the first-pass `Resource::ElementAttributes` check.

The `One { xml_space: true }` and `Many` arms remain on the old iterator. This
retains `xml:space` decoding and normalization, invalid-value errors,
duplicate-before-value behavior, the custom post-32 checked-attribute handoff,
and the second per-element limit check. Declarations call
`check_attribute_layout` directly and do not receive this summary.

## Error, limit, and state ordering

The candidate retains the existing order around each tag:

1. text-run checks, event/token/character checks, and `check_start` run first;
2. the complete manual attribute and namespace pass runs next;
3. the aggregate attribute charge (when applicable) occurs in
   `inspect_attributes`;
4. depth admission and the `Space` stack update follow.

Consequently an element-attribute limit still fails in the first pass at the
surplus attribute's `inner_offset + cursor`, before its name or value is read.
An aggregate limit on a valid zero/one non-`xml:space` tag still fails after
the manual pass at the tag-start offset, matching the old iterator path.
Namespace and malformed-input errors still precede that aggregate charge as
they did before. The candidate does not move any budget charge before the
manual pass.

The same state argument applies to all four call sites. Slice audits update
the observer counters after the fast charge; streaming audits use the same
absolute `start` offset and captured token; and source replacement windows use
the same `State::within`/`Counters` path. The summary reads no state that is
absent from a window replay, so it does not change window totals, namespace
scope, depth, or inherited-space handling.

## Counterexamples checked by inspection

The following classes remain outside the fast return or retain the old path:

* `<root a="1" b/>` and unterminated second values: the complete first pass
  returns its existing malformed or compactness error;
* duplicate names, including a duplicate malformed value: two attributes make
  the summary `Many`, so the checked iterator preserves duplicate precedence;
* `xml:space="maybe"` and escaped/non-normalizing values: the exact key keeps
  the old decoder and validation;
* leading `=`, unusual keys, whitespace around `=`, non-ASCII XML whitespace,
  and malformed separators: the manual first pass rejects or classifies them
  before a fast return can occur;
* namespace declarations and prefixed names: source namespace mutation and
  expanded-name validation remain in `check_start`;
* 31/32/33 and larger attribute lists: the first pass observes `Many` (or
  returns the existing per-element limit error), so the quick-xml/custom
  checked iterator and its 32-name handoff remain unchanged.

These observations are static. They are not substitutes for executing the
differential cases, especially because the quick-xml iterator intentionally
has lexical behavior that is not identical to the manual first pass.

## Required gates and residual risk

Before adoption, the candidate needs exact `Report` and `Error` parity for
source, compact, authored, and streaming entry points, including one-byte and
uneven `BufRead` chunks, BOM offsets, aggregate and per-element limits, UTF-8
failures, malformed tails, `xml:space`, namespace cases, and the 32/33-name
boundary. Source-replacement tests must also cover identical, complete,
window, original-failure, replacement-failure, and re-totaling paths.

The changed `check_start` result type and the extra summary propagation can
alter inlining, register allocation, or stack layout even if semantics are
identical. The 0905 quality receipts and the frozen paired native protocol
must therefore be evaluated separately. No native benefit, allocation result,
branch-coverage claim, or cross-format result follows from this review.

