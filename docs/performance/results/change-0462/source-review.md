# Change 0462 source review

## Scope and verdict

This review covers the archived frozen candidate diff against baseline
`dbd2f8ece` in:

- `crates/litchi-odp/src/codec/parser/codec/xml/validation.rs`
- `crates/litchi-odp/src/codec/parser/codec/xml/semantic.rs`

The candidate replaces repeated cached-prefix matching for shape attributes with
a private, bounded first-occurrence index. The source-level verdict is **ACK**:
the diff preserves the parser state machine and error behavior at the reviewed
call sites. Performance retention remains subject to the independent build,
assembly, correctness, and frozen matrix gates; this document does not turn a
source inspection into a latency claim.

The candidate source and binaries are retained in the 0462 evidence directory
for audit after the retention decision; this document does not require the
candidate to remain as production source.
The post-rejection working tree is baseline again, so the source inspection
below is anchored to the archived candidate text under
`source-artifacts/candidate/`, with the matching baseline text under
`source-artifacts/baseline/`.

## Implementation reviewed

`ShapeAttributeKey` is a private 17-member enum covering the union of the
presentation class/style/boolean fields, drawing name/style/layer/z-index/
transform fields, and the eight SVG geometry fields used by `shape_builder`.
`ShapeAttributeKey::target` supplies the exact namespace URI and local name for
typed calls. `from_target` and `from_cached` classify by URI and local name,
never by prefix or by the element namespace classifier.

`ShapeAttrs` wraps the unchanged `ElementAttrs` storage and carries a fixed
`[Option<usize>; 17]` first-occurrence array plus a clean-exhaustion bit. The
generic `AttributeIndex` policy keeps ordinary `ElementAttrs` calls on the
existing path through a zero-sized `NoAttributeIndex`; only the shape owner
uses the index. No public API, dependency, unsafe code, parser limit, or source
ownership boundary changes.

The typed `shape_builder` queries use `get_known`. A hit reads only the indexed
raw attribute and still performs the existing lazy value decode. A known key
without an index skips the cached linear scan only after the index says the key
is absent; if the shared iterator has only been partially consumed, the scan
continues. Clean iterator exhaustion is recorded separately from a malformed
attribute error.

## Semantic checks

- Namespace matching remains exact. `resolve_attribute` snapshots bound URI,
  unbound, and unknown-prefix outcomes in `ResolvedAttribute`; default
  namespaces therefore do not bind unqualified attributes, while aliases and
  element-local prefix rebinding resolve by URI as before.
- First occurrence is stable. `record` runs only after the raw attribute has
  been appended to `parsed`, and uses the first empty slot. Later duplicates do
  not replace an earlier cached occurrence.
- A freshly reached matching value is decoded before append. A decode failure
  advances the source iterator without creating an index entry, preserving the
  historical follow-up result. If a raw attribute was already cached while
  servicing another query, its indexed entry replays the same decode failure.
- An indexed hit is attempted before replaying `malformed`. Thus a valid
  cached value before a later malformed raw attribute still succeeds, while an
  uncached query that reaches that raw error reports the original XML error.
- Generic lookups still scan the cached prefix and then the shared iterator.
  They record recognized attributes as they pass them, so a later typed query
  can use the same first occurrence. A partial scan never proves absence.
- Drawing-attribute harvesting retains source order and its existing modeled
  attribute skips. The continuation appends and indexes a raw attribute only
  after harvesting and value decoding succeed, preserving the separate harvest
  error wording and failed-value reachability.
- The draw-style/presentation-style `Option::or` call remains eager. Both
  typed lookups are evaluated, so a later malformed attribute can still surface
  even when draw style already supplied a value.

## Independent focused evidence

The focused tests use `direct_lookup`, which iterates raw attributes and calls
the resolver directly without calling `ElementAttrs`, `ShapeAttrs`, or
`Parser::get_attr`. `oracle_drawing_attributes` independently reproduces the
historical fresh harvest. The frozen focused cases cover:

- all 17 typed keys and parity with the direct raw-attribute oracle;
- generic scanning followed by harvest and subsequent typed replay;
- namespace aliases, prefix rebinding, and the default-namespace rule;
- eager style fallback with a trailing malformed attribute;
- foreign-namespace invalid values being skipped by harvest;
- invalid harvested drawing values and the original harvest error wording;
- freshly reached versus previously cached invalid values;
- malformed XML and duplicate-attribute timing, including replay of the first
  cached value;
- the fixed index/layout diagnostic and the complete typed-key list.

## Receipts available at source review

The final warning-denied owner Clippy retry is recorded as PASS in
`checks/owner-clippy-r1.json`, and the release all-feature owner test receipt
is PASS in `checks/owner-tests.json`: 379 tests passed, with no failures. The
first owner-Clippy attempt exposed only integration issues in the initial
candidate (a missing namespace import, a stale unused import, and a test
assertion requiring equality for the crate error type); those were corrected
before the retry and are not part of the frozen source. The focused layout
diagnostic reports:

| Type | Size (bytes) |
| --- | ---: |
| `ElementAttrs` | 144 |
| `ShapeAttrs` | 424 |
| `ShapeAttributeIndex` | 280 |
| `ResolvedAttribute` | 80 |

The fixed index therefore adds 280 bytes of per-element state on this target.
That bounded stack cost is a review risk and must be considered with the tiny
input, peak/RSS, and allocation measurements; it is not evidence of a memory
benefit. The candidate build, final generated-code inspection, and frozen
performance matrix remain independent gates and are not inferred from these
correctness receipts. The candidate build, binding, baseline/candidate
assembly, comparison, counter, phase-diagnostic, and harness-test receipts are
PASS in the 0462 evidence directory. The harness test receipt records 387
passed, zero failed, and one ignored test.

## Manual generated-code review

The authenticated candidate and baseline release binaries were inspected
directly because the automated assembly heuristic only inspects direct calls
and mistakes their absence for scan removal. The heuristic's `eliminated` field is not
used as evidence here; the retained assembly receipts and symbol addresses are
authoritative for this paragraph.

The baseline `shape_builder` has 17 static call sites to
`ElementAttrs::get` and a `0x578` (1,400-byte) stack allocation. The candidate
has 17 corresponding static call sites to `ShapeAttrs::get_known` and a `0x698`
(1,688-byte) stack allocation. The extra 288 bytes in this function is
consistent with carrying the bounded index and typed lookup state through the
shape parse and is part of the practical memory review.

The candidate `ShapeAttrs::get_known` first selects the enum slot, checks its
first-occurrence discriminant, loads the indexed `parsed` entry, and calls
`ElementAttrs::lookup` once. The return path handles the existing value decode
result before returning. When the slot is `Known(None)`, the assembly checks
the stored malformed state and clean-exhaustion state before entering the
shared raw-iterator path. A cached-prefix loop remains as a defensive fallback
if an indexed `Some` entry ever fails the cached-key predicate; the hot valid
indexed path does not enter that loop. The raw continuation still performs the
existing resolve, lazy lookup, append, and index-record sequence.

The generic candidate `ElementAttrs::get` retains the old cached-prefix and
raw-iterator mechanism, with a 328-byte frame versus 328 bytes in the
baseline; the inspected instruction counts are 283 candidate versus 285
baseline. This confirms that the fixed index policy is confined to the shape
owner and that generic attribute users do not acquire the typed index state.
These observations establish the intended mechanism only. They do not predict
end-to-end latency or offset the measured 280-byte per-element index cost.

## Measured gate and review conclusion

The frozen primary comparison contains 24 reports and 720 retained samples.
The protocol requires at least a 3% normal-lane p50 improvement for both
medium and large shapes in both repeats, with each independent bootstrap
median-delta interval's upper endpoint below zero. The medium lanes pass that
threshold: R1 is -3.4649% (95% upper endpoint -3.3210%) and R2 is -3.1917%
(upper endpoint -2.6071%). The large lanes are consistently faster but remain
below the required threshold: R1 is -2.6602% (upper endpoint -2.2252%) and R2
is -2.6279% (upper endpoint -2.3899%). Therefore
`predeclared_latency_gate_pass` is `false`; the candidate does not satisfy the
retention gate.

All allocator lanes report equal allocation bytes, allocation calls,
deallocation calls, peak-above-entry bytes, reallocations, and retained-live
bytes. The comparison reports no adverse greater-than-5% flags and no
allocation-increase review flags. Whole-process counter deltas are
instructions -3.3494%, cycles -3.0144%, branches -2.7522%, cache misses
+1.2504%, branch misses +1.2247%, page faults +2.1945%, and context switches
-2.4540%; these include setup, warmups, checks, and reporting. The
supplementary phase profile shows large-shape p50 changes of -3.6631% and
-4.1715% for commit, and -3.9477% and -5.6825% for snapshot-open, in R1 and
R2 respectively; those clocks are diagnostic and do not override the primary
gate.

The source-level behavior is therefore ACKed for correctness, while the
candidate is **rejected for retention** under the frozen protocol. The
280-byte per-element index and 288-byte `shape_builder` frame increase have
no separately quantified practical memory benefit that could justify the
failed large-shape threshold. Both production files are restored byte-exact to
baseline revision `dbd2f8ece`. This review records the archived rejected
candidate source, correctness, assembly and measurements. Retained source and
receipts remain audit artifacts after the staged executables are cleaned.
