# XML streaming workspace review

Status: independent source review accepted for the 0483 XML workspace change.
The reviewed production diff is limited to the duplicate-attribute scratch term
in `Limits::streaming_memory_upper_bound` and its documentation/tests. Root's
recorded `xml-audit-test-dev-01` and `xml-tests-dev-02` receipts both completed
with exit code zero, including the new aggregate-versus-token workspace cases.
No Cargo command was run for this review.

This review checks the bound against the locked `quick-xml 0.41.0` source and
the current `GuardedBufRead` implementation. It treats the public helper as a
checked implementation envelope. Caller-owned `BufRead` storage, allocator
metadata and rounding, fixed parser values, and formatted error strings remain
outside the stated contract, as the API documentation says.

## Verdict

The changed attribute term is conservative under the current parser and guard:

```text
T = max_token_bytes + 1
P = min(max_attributes, T)
E = P + 1
```

`E` is the maximum key-entry count charged to quick-xml's per-event duplicate
scratch, including one candidate that can be inserted before the auditor
reports an aggregate attribute-limit error. The aggregate `state.attributes`
counter still uses the unmodified `max_attributes` limit, so a document may
use that budget across many small events. The new cap affects only one
lexical event's temporary vectors and hash table.

The helper remains a dependency-versioned size envelope. If quick-xml changes
its duplicate-check data structures, or if this crate starts using a different
attribute iterator, the range/hash factors must be re-audited.

## Why the first rejected attribute is covered

`BytesStart::attributes()` borrows the current parser event and creates an
`Attributes` iterator. Its `IterState` owns a `Vec<Range<usize>>` called
`keys`; after the small-tag threshold it also owns a `HashSet<u64>` prefilter.
The key ranges point into the event bytes, so the iterator does not retain a
second copy of the tag, but the vectors are real per-event allocations.

The order in the current audit is significant:

1. `tag.attributes().next()` parses an attribute and runs quick-xml's
   duplicate check.
2. For a unique key, `IterState::check_for_duplicates` pushes its range into
   `keys` before returning the `Attribute`.
3. `inspect_attributes` then increments the document-wide counter with
   `checked_add` and can return `Error::Limit`.

Thus a profile with `max_attributes = A` can have `A` accepted attributes in
the event and one additional unique key already present when the `(A + 1)`th
attribute is rejected. The same ordering can leave one key recorded before an
unquoted or otherwise malformed value is reported. `E = min(A, T) + 1` covers
both cases. A duplicate-name error does not add a key; in the hashed path its
hash insertion can only reuse existing table capacity, so it does not require
another count beyond `E`.

For example, at the 32-attribute threshold, the first candidate after 32
unique keys creates quick-xml's hash prefilter, inserts the candidate hash,
pushes the candidate range, and only then lets the audit counter reject it.
The `+1` is therefore needed on the same path that first allocates the hash
table. Returning on the first iterator or aggregate error prevents any later
attribute from accumulating scratch.

The event itself is bounded by the token window. Even on a malformed event,
each iterator step advances through the bounded event bytes; the deliberately
loose `T` candidate bound is therefore safe for empty-key and truncated-value
parser paths as well. A complete event whose raw length is `max_token_bytes +
1` is rejected by the audit's token check before `inspect_attributes`; charging
that sentinel-sized case still makes the workspace helper conservative.

## Range and hash capacity factors

The current helper computes:

```text
attribute_ranges = max(4 * size_of::<Range<usize>>(), 2 * E * size_of::<Range<usize>>())
attribute_hashes = max(8 * H, 4 * E * H)
H = size_of::<u64>() + size_of::<u8>()
```

The range factor covers the pinned `Vec`'s geometric growth from the empty
iterator through the first rejected key. The fourfold hash factor covers the
initial `HashSet::with_capacity_and_hasher(keys.len() * 2)` at the 32-key
transition, a subsequent table growth envelope, and one 64-bit value plus one
control byte per bucket. The minimum floors cover the small-profile case where
the configured event contains too few attributes to reach a normal geometric
capacity. These factors are intentionally larger than the logical entry count;
they are capacity envelopes rather than serialized data sizes.

The table and vector objects themselves live inline in quick-xml's iterator;
their fixed metadata is covered by the helper's fixed parser-value convention.
Allocator bookkeeping and implementation-specific alignment remain outside the
documented envelope. The checked arithmetic in the public helper fails closed
if any term or sum cannot be represented by `usize`.

## Why the token and byte guards make the premise hold

Quick-xml's buffered source methods append every visible `fill_buf()` slice to
the supplied event `Vec` while searching for a markup delimiter. Tiny source
chunks alone would not bound that vector. The current guard is the relevant
pre-growth control:

* `fill_buf()` limits each exposed slice by both `max_total - total` and
  `(max_token_bytes + 1) - token`.
* `consume()` copies only the exposed amount into `captured`, consumes the same
  amount from the caller source, and increments both counters.
* Once the lookahead window is exhausted, `fill_buf()` returns a typed window
  error instead of exposing another byte. An underlying total-byte overrun is
  refused by the same method before quick-xml can append it.
* The event buffer, `captured`, and `exposed` are each reserved for `T`, so
  quick-xml's `read_text`, `read_ref`, `read_with`, and `read_bang_element`
  paths cannot append an unbounded hostile text, reference, tag, PI, comment,
  CDATA, or declaration token.

The guard also preserves the three-byte BOM prefix outside these vectors. A
BOM event may use the reusable `bom_raw` window, and the helper includes that
optional `T` allocation in its six token capacities. The fixed prefix,
history, and combine arrays are charged separately by the existing 12-byte
term.

The parser records an open element name before the audit depth check. The
existing `(max_depth + 1)` open-name and index terms therefore remain required;
the attribute cap does not weaken that depth-plus-one accounting. Attribute
scratch is released with the `Attributes` iterator after each event, while the
open-name and `xml:space` stacks remain document state.

## Limits and exclusions

The source-derived cap does not change acceptance policy. An input with one
attribute on each of many elements can consume the full aggregate attribute
budget while keeping each event's duplicate scratch small; the dedicated
aggregate test in the XML audit receipt confirms that behavior. Conversely, a
single large start tag is constrained by both the aggregate budget and the
token window before quick-xml can retain an unbounded iterator state.

This review does not turn the helper into a complete RSS guarantee. It excludes
the caller's own buffered storage, allocator retention and metadata, fixed
parser scalars, and error-detail strings. It also relies on the current
quick-xml hash-table layout and geometric capacity behavior. Those exclusions
are explicit and should remain attached to any caller-facing memory claim.

## Evidence

Root-owned receipt `validation/xml-audit-test-dev-01.stdout` reports 11 passed
unit tests, including:

* aggregate attribute accounting independent of token scratch;
* equality of the large aggregate-attribute and token-sized scratch profiles;
* checked arithmetic overflow returning `None`; and
* token-window rejection before unbounded parser growth.

Receipt `validation/xml-tests-dev-02.stdout` reports the complete
`xml-minifier` package test run passed. The static source review above is the
evidence for the first-rejected-attribute allocation ordering; no additional
build or runtime command was run here.
