# 0800 source review: key-prefix correctness gate

This is a source-only review of the 0800 packet. I inspected the current
`litchi-opc` helper, the five archived helper copies, the 0.41.0 quick-xml
iterator, the inherited 0799 semantic policy, and the retained 0800 edge probe.
I did not edit production, run a build or test, or run a native or profiler
capture.

## Disposition

The 0800 no-replay design remains a deferred performance candidate. The minimal
correction is now present in all five `correction/after` helper copies. The
`deferred-candidate/` archive is intentionally frozen against the
pre-correction base; it must be rebased onto `correction/after` and retested
before anyone builds or measures it. Do not edit that archive incidentally
while preserving its source-frozen record.

The repair is independent of the no-replay state-machine hypothesis. It should
be gated first with the same quick-xml differential oracle and malformed-input
checks. The 0799 39-case preflight and its fail-closed decision rule remain
frozen for any later no-replay measurement: both `distinct-1/consume` and
`distinct-2/consume` need the stated benefit, the 18 protected rows retain
their veto, all other regressions remain visible review triggers, and no result
authorizes production adoption.

## Defect and exact parser rule

The existing helper finds the first non-whitespace byte at `start`, then scans
from that same byte for `=` or XML whitespace. quick-xml's `IterState::next`
has already consumed the byte returned by its `find` for `start_key`; its next
search begins at `start + 1`. Therefore a raw prefix beginning with `=` is a
key beginning with `=`, rather than an empty key.

The retained edge probe demonstrates the mismatch. For a prefix of 32 or more
accepted attributes followed by ` =n0="unterminated`, quick-xml reports
`Duplicated(position, first_position)` before scanning the malformed value,
while the current bounded helper reports `ExpectedQuote`. The 32-attribute
case is `Duplicated(302, 2)` in the retained probe; the longer cases differ at
their corresponding positions. The one- and 31-attribute rows do not reach
the affected `OwnCheck` path, which is why they pass and do not disprove the
bug.

`name_at` must mirror the parser's consumed-byte boundary:

```rust
let rest = tag.get(from..)?;
let start = from + rest.iter().position(|byte| !is_whitespace(*byte))?;
let after_first = tag.get(start + 1..)?;
let length = after_first
    .iter()
    .position(|byte| *byte == b'=' || is_whitespace(*byte))
    .map_or(after_first.len(), |length| length);
let end = start + 1 + length;
Some((start, &tag[start..end]))
```

The offset arithmetic is bounded by the slices it indexes: `start` is an
existing byte, so `start + 1` is at most `tag.len()`, and `after_first` is the
remaining suffix, so adding its selected length returns an end no later than
`tag.len()`. The important invariant is that the first non-whitespace byte is
included in the returned name and excluded from the delimiter search. The
helper remains a borrowed, single-pass prefix scan with no allocation.

This exact change belongs in all five canonical copies:

* `litchi-opc`;
* `litchi-ole-common`;
* `litchi-sign`;
* `litchi-xldm`; and
* `xml-minifier`.

The public or crate-private visibility boundaries, the OOXML re-export, and
the unchecked `first_wins` and `unchecked_attributes` alternatives do not
change.

## Semantic conditions for the correction

The repaired helper is used only while translating an already returned
unchecked-parser error in `duplicate_before_value`. The guard must continue to
remap only `UnquotedValue`, `ExpectedValue`, and `ExpectedQuote`. In
particular, `ExpectedEq` must remain `ExpectedEq`; a repeated raw prefix that
has no recognized `=` has not reached quick-xml's duplicate check.

The correction must preserve these boundaries:

* XML whitespace remains exactly space, tab, carriage return, and line feed.
  Other bytes must not be silently treated as separators.
* A key beginning with `=` is compared as that full raw key prefix, with the
  duplicate position at its first byte and the first position stored from the
  earlier raw key.
* Whitespace between a key and `=` still qualifies for duplicate-before-value;
  whitespace followed by another non-whitespace byte still yields
  `ExpectedEq`.
* A duplicate followed by an unquoted, missing, or unterminated value must
  retain quick-xml's `Duplicated` error priority and positions. The existing
  bounded fallback obtains the lexical error by asking the unchecked parser
  for the item and then translates it; this correction does not claim to avoid
  scanning the value. The deferred candidate's raw `key_at` preflight is the
  separate design that can refuse before that parser call. A nonduplicate
  malformed value must retain its original lexical error and position.
* Valid values, borrowed key/value slices, first-error ordering, and fused
  exhaustion are unchanged.

The helper mirrors quick-xml's raw prefix grammar and does not validate XML
`Name` syntax. Tests must not impose an XML-name validator on this layer. For
example, in a raw prefix beginning `==n0=`, quick-xml consumes the first `=`
as the first key byte and recognizes the second `=` as the delimiter; the
following `n0` is then an unquoted value and the resulting lexical error has
priority. That behavior is part of the compatibility contract for malformed
`BytesStart` content.

The direct semantic cases should include a leading-`=` key and a mixed pair,
such as `=n0="1" =n0="unterminated` and `=n0="1" = "unterminated`, in
addition to the retained 32/33/34-prefix edge. The mixed form protects against
accidentally treating the first `=` as a delimiter and thereby confusing the
key `=n0` with an empty key. Cases with a duplicate key lacking `=` must remain
`ExpectedEq`, and ordinary keys with spaces around `=` must retain their exact
positions.

## Deferred no-replay candidate review

The archived state machine has a coherent shape once rebased onto the
correction: `First` and `Second(offset)` use an unchecked parser;
`ShortTwo { first_end, next_start }` checks the third raw prefix before asking
the parser for its value; and a successful third item seeds an owned backend
from borrowed key prefixes without reparsing either earlier value. The
`deferred-candidate/` source is deliberately still pre-correction, so this is
a design assessment only. The rebased precheck must use the same consumed-byte
`key_at` rule as quick-xml, including the leading-`=` case, and must compare
only when the prefix has a recognized `=`.

The transition and layout still require direct gates when the candidate is
resumed:

* `ShortTwo` carries two offsets, so `size_of::<CheckedAttributes>()` must be
  measured rather than inferred from the old `QuickXml(usize)` layout. Clone
  must preserve both offsets and the underlying unchecked iterator at every
  boundary.
* A third duplicate must be refused before its value is parsed. A third
  nonduplicate lexical error must preserve its original error; a third
  success must be inserted exactly once and must not replay the first two.
* An owned vector backend may remain through at most 32 names and then switch
  to the ordered map, or the map may be selected earlier. Either choice must
  retain a hostile-tag bound of `O(n log n)` or better and must not permit a
  duplicate during the transition. The map's first-position entry must be the
  earliest raw key position.
* `OwnCheck::from_prefix` should fail closed if a supposedly successful prefix
  cannot be recovered. A release build must not rely on a debug assertion to
  maintain name-count or key-presence invariants for untrusted bytes.
* The map precheck must run before `attributes.next()`. If it misses a duplicate
  and the later `check_attribute` catches it only after parsing the value, the
  duplicate-before-value contract is broken even when the final error variant
  looks correct.

The correction itself does not alter valid-value behavior or asymptotic work;
it only makes the existing error translation agree with quick-xml at the
leading-`=` boundary. No timing, allocation, workflow, or production-adoption
claim follows from this source review.
