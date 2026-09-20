# Change 0701 design — outlined first-byte MCE marker search

Status: static design record captured before measurement. No production
source, Cargo manifest, test, or measurement file was changed by this
document. At design capture, the candidate was an unmeasured hypothesis; the
implemented result and final retention review are recorded in
[`code-review.md`](code-review.md).

## Target and reason for a new experiment

The shared buffered MCE entry point in
`crates/litchi-ooxml-common/src/mce/codec.rs` currently decides whether to
enter the XML processor with a fixed-size byte-window search:

```rust
if !xml
    .windows(NAMESPACE.len())
    .any(|window| window == NAMESPACE.as_bytes())
{
    // preserve the borrowed, marker-free result
}
```

Change 0700 replaced that predicate with the existing safe
`memchr::memmem::find` search. It removed work from long marker-free buffers,
but the library searcher's per-call setup made tiny marked inputs about 45%
slower (roughly 210 ns), and the repeated real edit and no-op workflows were
about 1.5–3.4% slower. The direct candidate was rejected and the production
codec was restored. Those results motivate a distinct, smaller search
hypothesis; they do not predict a 0701 gain.

The 0701 candidate keeps the dispatch boundary and the existing parser path,
while using `memchr` only to skip bytes that cannot begin the fixed URI. An
exact `starts_with` check handles each possible first-byte position. The
helper is private and outlined with `#[inline(never)]` so the searcher's local
state cannot be pulled into `process_markup_compatibility`'s frame by
inlining. The attribute is a code-layout hypothesis, not a performance claim.

## Smallest intended source change

The candidate should add one private helper and replace only the early
predicate:

```rust
#[inline(never)]
fn contains_mce_namespace(mut remaining: &[u8]) -> bool {
    let needle = NAMESPACE.as_bytes();
    let first = needle[0]; // NAMESPACE is a non-empty, fixed URI constant.

    while remaining.len() >= needle.len() {
        let valid_starts = remaining.len() - needle.len() + 1;
        let Some(offset) = memchr::memchr(first, &remaining[..valid_starts]) else {
            return false;
        };
        if remaining[offset..].starts_with(needle) {
            return true;
        }
        remaining = &remaining[offset + 1..];
    }
    false
}
```

The production call site should become:

```rust
if !contains_mce_namespace(xml) {
```

The final implementation may use an equivalent checked retrieval of the
first byte if the reviewer prefers to make the non-empty-constant invariant
explicit in code. It must retain the same loop and valid-start bound. There
must be no new static cache, `OnceLock`, global state, dependency, unsafe
block, hand-written SIMD, or replacement of the existing `find_bytes` helper.
`find_bytes` is also used by active-offset marker processing; changing that
separate path would enlarge the hypothesis and invalidate this design.

The intended source diff therefore has these boundaries:

- `NAMESPACE` and its byte representation remain the existing constants;
- `memchr` remains the already direct, locked crate dependency of
  `litchi-ooxml-common`;
- the input-limit check stays before the helper call;
- the marker-free output-limit check, borrowed `Cow`, and default `Report`
  stay exactly where they are;
- every exact URI occurrence, including one in text, a comment, or malformed
  bytes, still selects the existing `Reader` path;
- `Reader`, `BoundedOutput`, `Frame`, `Ctx`, namespace ownership, error order,
  output limits, streaming MCE, active-offset selection, and public APIs are
  untouched.

No test or production result is inferred from this source sketch. The
candidate source must be frozen and reviewed after formatting before any
timing or allocation run.

## Correctness argument

Let `m = needle.len()`. The MCE URI is non-empty, so `m > 0`. At the top of
each loop, `remaining.len() >= m`. The expression
`remaining.len() - m + 1` therefore cannot underflow and is the number of
valid starting offsets in `remaining`. Searching only
`remaining[..valid_starts]` includes the first and last complete windows and
excludes every trailing position that cannot contain the complete URI.

`memchr` returns an offset whose byte equals the URI's first byte. The slice
starting there has at least `m` bytes because the offset is within the valid
start prefix. `starts_with(needle)` is therefore safe and is exactly the
remaining byte comparison performed by the old window predicate. A failed
comparison advances to `offset + 1`, so the next iteration examines every
later possible start, including overlapping candidates. The offset is always
at least zero and the remaining slice strictly shrinks, so the loop
terminates.

This proves boolean equivalence for all byte slices, not only UTF-8 XML:

- empty and shorter-than-URI input returns `false` without a subtraction or
  slice;
- an exact URI at offset zero is found on the first iteration;
- an exact URI beginning at the final valid window is included by the `+ 1`
  bound;
- a URI followed by arbitrary bytes, including non-UTF-8 bytes, is detected
  lexically;
- first-, middle-, and final-byte near misses fail the exact check and then
  continue at the next byte;
- repeated and overlapping first-byte prefixes cannot hide a later match;
- an exact URI in text or a comment still selects the parser, and an exact
  URI in malformed input still reaches the parser's existing XML error path;
- a first-byte occurrence in the final fewer-than-`m` bytes is correctly
  ignored because it is not a valid start.

The helper returns only a boolean and performs no allocation or fallible
operation. It cannot change typed errors itself. All error, output, report,
and ownership behavior remains determined by the pre-existing branch after
the predicate.

## Cost model and risks

The helper uses O(1) explicit storage. Let `n` be the input length, `m` the
fixed URI length (currently 59 bytes), and `k` the number of candidate
first-byte positions. The `memchr` scans and slice advances cover the input
once in logical terms; each candidate can perform one comparison of at most
`m` bytes. The bounded cost is therefore `O(n + k*m)`, or `O(n*m)` in the
worst case. Since `m` is a fixed protocol constant, this is linear in input
length with a potentially material constant.

The favorable case is a marker-free buffer with few or no bytes equal to the
URI's first byte. `memchr` skips those regions in one search operation, and
the helper avoids constructing the general-purpose substring searcher that
hurt 0700's tiny marked controls. The unfavorable case is an adversarial
buffer containing the first byte at nearly every valid position, especially
repeated near-prefixes that differ late in the URI. Such a buffer performs a
long exact comparison at each position. The 0700 repeated-prefix control must
be rerun against 0701; its prior result cannot be transferred to a different
algorithm.

Outlining the helper has two opposing effects. It can keep the large search
implementation and its locals out of the caller's frame, but it adds a call
boundary to every invocation and may prevent caller-specific optimization.
Tiny marker-free, tiny marked, and real MCE-positive inputs are consequently
first-class controls. The candidate must not be kept based on assembly size,
one synthetic scan, or a favorable single process leg.

There is no new retained memory, cache, lock, execution context, source
provider, or resource budget. The existing input and output limits remain the
only admission boundaries. Generated code size, helper frame size, native
text, and processor stack reservation still require measurement because
`#[inline(never)]` controls source inlining but does not guarantee a smaller
binary or stack.

## Evidence required before retention

The 0701 packet should reuse the 0700 harnesses, corpus, and acceptance
invariants as fresh runs against 0701 source and binary identities. Reusing a
driver or corpus means re-running it with new receipts; it does not make the
0700 measurements evidence for this candidate. The sequence is:

1. Capture the restored 0701 baseline, including the source census, focused
   tests, build inputs, lockfiles, binary identities, and native/profile
   environment.
2. Apply only the helper and early-predicate change, format it, freeze the
   candidate source diff, and run the focused MCE tests against both source
   states.
3. Require exact differential parity for the existing 192-case, five-profile
   oracle: output bytes and lengths, `Cow` ownership, complete `Report`, and
   typed/debug error identity must agree for both binaries.
4. Re-run the marker-control family, the 13-workflow native matrix, the
   10-case refusal matrix, the independent allocation comparison, hardware
   profile, assembly/code-size inspection, and all existing integration and
   evidence gates with fresh hashes.
5. After those gates are terminal, run the isolated longer follow-up with the
   six native cases and complete ten-case refusal matrix. Preserve all raw
   distributions and every review trigger.

The focused semantic matrix must demonstrate the edge cases above, either
through retained 0700 focused tests plus the oracle/control corpus or through
new focused witnesses if a case is not directly represented. In particular,
the final review must be able to point to evidence for offset zero, the final
valid window, first/middle/final near misses, overlap, arbitrary bytes,
text/comment placement, malformed XML, and input-before-output limit order.

## Retention and rejection rule

Retention requires a statistically credible and practically useful benefit
on representative shared OOXML work, with no unexplained correctness or
resource cost. Every individual scenario remains visible. The initial review
trigger from [`docs/GOAL.md`](../../../GOAL.md) applies: any approximately
greater-than-5% latency or throughput regression, 5% peak-RSS regression, or
material scaling loss is a review trigger and cannot be hidden by a geometric
mean.

The candidate should be rejected if the lower-setup helper still produces
recurring regressions on the real edit/no-op workflows, if tiny marked or
MCE-positive work pays a material setup cost without a representative offset,
if the repeated-prefix cost is unacceptable, or if code/stack growth has no
compensating end-to-end benefit. Rejection preserves the measured candidate
witness and restores only the production codec; retained focused tests and
the evidence record may remain when they document the proven lexical
boundary. No production performance claim is allowed until the final review
passes these conditions.

## ADR and goal constraints

The design follows the accepted records consulted before writing it:

- [ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR 0002](../../../adr/0002-crate-topology.md), and [ADR 0024](../../../adr/0024-current-topology.md): the helper stays private to the existing shared OOXML owner, introduces no public type or dependency edge, and does not expose archive state.
- [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md): the codec is a read/preprocessing dispatch detail; immutable snapshots, edit publication, exact no-ops, patches, and conflicts are unchanged.
- [ADR 0005](../../../adr/0005-io-memory-and-performance.md) and [ADR 0031](../../../adr/0031-execution-context-budgets.md): the scan is bounded by the existing source limit, allocates no retained state, adds no cache or worker, and uses no ambient execution or I/O.
- [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md): the search remains byte-oriented and lexical; it does not validate, repair, normalize, or reinterpret malformed or unknown markup, and the existing parser/refusal boundary remains authoritative.
- [ADR 0008](../../../adr/0008-migration-and-verification.md): baseline/candidate source identities, differential preservation evidence, adversarial controls, and full gates are required before any support or performance claim.
- [ADR 0010](../../../adr/0010-facade-archive-ownership.md) and [ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md): no facade, archive, physical package, or relationship ownership moves.

No accepted ADR is amended or proposed by this experiment. No unsafe code,
custom SIMD, global state, new dependency, API, limit, or serialization rule
is part of the hypothesis.
