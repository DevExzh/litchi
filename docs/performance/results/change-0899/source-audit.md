# 0899 source audit: allocation-free first-prefix duplicate checking

**Status: read-only source audit; no production candidate or adoption decision.**

This audit evaluates whether the next duplicate-attribute experiment should
remove quick-xml's first small-tag duplicate-name allocation. It does not edit
Rust source, build, test, run a workload, or run a profiler. The production
tree is the restored `4f8f667983` head after the 0898 rejection. The accepted
architecture remains bound to `506aa2edae`; the current commit history and
working tree were checked before this report was written. The two unrelated
untracked files already in the workspace and the other 0899 packet artifacts
are left untouched.

## Evidence and scope

The source review covered:

* [`docs/GOAL.md`](../../../GOAL.md), the accepted ADR index and the accepted
  records governing measured evidence, crate ownership, malformed-input
  handling, limits, and public API boundaries;
* the canonical helper at
  [`crates/litchi-opc/src/xml_attributes.rs`](../../../../crates/litchi-opc/src/xml_attributes.rs),
  its shared tests, and the four canonical copies in `litchi-ole-common`,
  `litchi-sign`, `litchi-xldm`, and `xml-minifier`;
* quick-xml `0.41.0` at
  `/home/zhuhe/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/quick-xml-0.41.0/`,
  especially `events/attributes.rs`; and
* the retained 0764, 0794–0806, and 0897 records.

The applicable accepted constraints are ADR 0001's strict public layers, ADR
0002's downward crate topology, ADR 0005's measured-performance and budget
rules, ADR 0006's preservation and malformed-input defenses, ADR 0008's
verification gates, ADRs 0010/0011/0024's current ownership topology, and ADR
0032's limits on retained derived state.

The 0897 retained profile pair records 673,663 whole-process allocation events
per process and 518,660 events in the scan closure over 20 scans. The
`quick_xml::events::attributes::IterState::check_for_duplicates` stack group
contains 85,560 events, or 4,278 per scan. That is stack-presence allocation
evidence for a path, not proof that all of those calls are removable by this
helper or that removing them will lower latency.

The fresh native leaf preliminary ranges supplied for this audit put duplicate
self cost at only 0.7515–0.8327% of the whole process, the reserve region at
0.9685–1.0827%, and codec `start` at 12.99–13.50%. Those numbers are a
prioritization signal, not a retained 0899 measurement or a latency claim.
The allocation hypothesis is therefore small in the end-to-end budget even
though the call count is real.

## What the current code does

The wrapper deliberately has two phases:

| prefix | current path | source consequence |
|---|---|---|
| first 32 successful items | `tag.attributes()` with quick-xml checks enabled | quick-xml parses the item and pushes its key `Range<usize>` into `IterState.keys`; the first successful key can allocate a `Vec` |
| the 33rd request and later | checks are disabled for the current item; the wrapper then replays the first 32 names through `unchecked_attributes()`, seeds `BTreeMap<Name<'a>, usize>`, and boxes `OwnCheck` | the replay avoids quick-xml's unkeyed hash path, but pays map construction and a second lexical walk of the prefix |

`quick_xml::events::attributes::IterState` contains a `Vec<Range<usize>>`
even though the empty iterator itself is allocation-free. `check_for_duplicates`
pushes a range after the key and its `=` have been recognized, before the value
is read. At 32 names quick-xml switches to a hash prefilter; the wrapper turns
that check off and uses its ordered map instead. This is the bounded strategy
adopted after the 0764 hostile-input work.

The public `CheckedAttributes` contract is stronger than “eventually reject a
duplicate.” It returns the same borrowed `Attribute` items and first error as
quick-xml, sets `Done` after an error or exhaustion, implements `Clone`,
`Debug`, and `FusedIterator`, and preserves raw byte positions. `end_of`
relies on quick-xml returning a borrowed value and advances from the closing
quote only after a successful item. The current `name_at` correction follows
quick-xml's consumed-first-byte rule, so a leading `=` belongs to the key.

The helper is copied rather than dependency-shared in five owners. The OPC and
OLE common `BytesStartExt` trait and `CheckedAttributes` type are public; in
particular, `litchi-crypto` imports the OLE helper. The OPC helper is re-exported
through OOXML common. The formula `first_wins` helper is a separate lenient
contract and is out of scope.

## Exact behavior a candidate must preserve

The key is the raw byte slice, including any prefix. Namespace equivalence or
decoded values must not affect duplicate identity. Quick-xml's order is:

1. skip exactly space, tab, carriage return, and line feed;
2. consume the first non-whitespace byte as part of the key;
3. recognize a direct or separated `=`;
4. check the raw key for a duplicate; and
5. only then parse the value.

This gives the following required outcomes:

| input condition at the next key | required result |
|---|---|
| key has no recognized `=` | delegate to quick-xml and retain `ExpectedEq`; do not turn a repeated raw prefix into a duplicate |
| key is a repeated raw name and `=` is recognized | `Duplicated(current_key_start, first_key_start)` before scanning its value, then `Done` |
| unique key with `UnquotedValue`, `ExpectedValue`, or `ExpectedQuote` | retain the quick-xml error and byte position, then `Done` |
| no more non-whitespace bytes | `None`, then fused `None` |

The malformed duplicate cases are essential. A valid duplicate, an unquoted
duplicate, a missing-value duplicate, and a long unterminated duplicate all
must report the same duplicate position and must not scan the duplicate value.
The current 0800 correction is also essential: for a prefix such as `=n0`,
the first `=` is key data and the later delimiter is the recognized equals
sign. A preflight that starts scanning at the first byte again will report
`ExpectedQuote` or `ExpectedEq` instead of `Duplicated`.

The iterator stops after the first error, so an internal unchecked parser
cursor does not need to expose quick-xml's recovery after an error. It must,
however, be set to the terminal state before returning that error. A clone made
before an item, at the prefix-to-map handoff, or after an error must produce
the same remaining sequence and fused exhaustion. Derived `Debug` output is
observable on a public value, so a layout/state change should be recorded and
covered even though the fields are private.

## Concrete candidate shape

The viable source-level mechanism is a **borrowed raw-key prefix index**:

* construct quick-xml's `Attributes` with duplicate checks disabled;
* keep an inline count and at most 32 raw key starts, for example a safe
  `count: u8` plus `[usize; 32]` of starts; do not retain `Attribute` values,
  `Cow`s, decoded names, or copied name bytes;
* before every unchecked parser call, run a `key_at`-style preflight from the
  independent `next_start` cursor. It must return the key span and whether a
  recognized `=` exists, using the consumed-first-byte rule above;
* for a recognized `=`, compare the current raw slice with the bounded stored
  names. A start-only table can derive each stored key's end by scanning only
  from that key start to its first `=` or XML whitespace; a separate
  `start/end` table is a valid diagnostic variant if that scan costs more;
* on a duplicate, return `Duplicated` immediately and set `Done`; on a key
  without `=`, call quick-xml so `ExpectedEq` remains authoritative; and on a
  unique key, call the unchecked parser and record its start only after `Ok`;
* when the 33rd request begins, derive the 32 borrowed names from those starts
  and seed the existing ordered `BTreeMap<Name<'a>, usize>` while keeping the
  unchecked parser cursor. Do not reconstruct a checked iterator or replay the
  first 32 values. Preflight the current 33rd key before asking the parser for
  its value; and
* in the ordered phase, perform the same raw preflight before every parser
  call, insert only successful unique names, and leave `next_start` unchanged
  on errors.

This differs from the rejected 0794 four-name/vector design and from 0887's
inline attribute storage: only bounded byte positions are retained, and
values are never stored in the duplicate index. It also avoids an unkeyed hash
table. A full fixed array is still a deliberate layout cost, not a free
optimization. On a 64-bit target, 32 `usize` starts consume 256 bytes before
the count and cursor; caching both start and end consumes roughly 512 bytes.
The current iterator is 120 bytes. An inline `Name` array is also roughly 512
bytes and has already appeared in the rejected family of designs. The
start-only form is the smallest safe representation that preserves arbitrary
`usize` byte positions without imposing a new public tag-length limit.

The array can be initialized lazily on the first non-empty request to avoid
unnecessary element stores on an exact-empty path, but laziness does not reduce
the iterator's layout. `MaybeUninit` or packed `u32` offsets would either add
unsafe invariants or silently cap public byte positions; neither is justified
by the present evidence. A new `smallvec` or `arrayvec` dependency would also
need to be replicated across five owners and is outside the current crate and
manifest boundaries.

The transition is therefore:

| successful prefix count | state before the next request | work before parsing a value |
|---:|---|---|
| 0–31 | inline raw-key index | bounded direct byte comparisons; store the new start only after a successful item |
| 32 | inline raw-key index | seed the existing ordered map from 32 borrowed spans, then preflight the 33rd key |
| 33+ | ordered map | ordered duplicate lookup, then unchecked lexical parse; insert only a successful unique key |

The candidate must not use the simpler “parse unchecked, then remap the
error” approach for duplicate values. That is the 0794 shape and is exactly
why its long quoted and unterminated duplicate rows scanned the whole value.
The raw preflight must happen before `attributes.next()` at every phase.

## Bounds, hostile input, and API risks

The first prefix has a fixed maximum of 32 names. Direct comparisons are at
most 32 per request; with start-only spans, key-end discovery adds bounded
scans of key bytes but never scans an earlier value. At the handoff, the
ordered map retains the current `O(n log n)` comparison behavior and avoids
quick-xml's unkeyed hash prefilter. No new per-tag allocation is introduced
for zero through 32 successful names, but tags over 32 still allocate the
existing map and box state. The helper itself has no caller-limit parameter,
so this is a bounded duplicate-check algorithm, not a replacement for the
per-element and XML-byte ceilings enforced by its callers.

The principal risks are:

* **layout and movement:** every `CheckedAttributes` construction, move, clone,
  and drop carries the inline array, even for zero or one attribute. This is
  the same class of cost that caused the 0794 candidate to grow from 120 to
  192 bytes; the proposed starts-only layout is smaller than a name-span or
  `Attribute` array but still needs a direct size/stack diagnostic;
* **small-prefix instruction cost:** each accepted item adds raw key
  preflight, position comparison, and `end_of` bookkeeping that the current
  first 32 items leave to quick-xml. A saved allocator call can be outweighed
  by these stores and branches, especially when the duplicate leaf is below
  one percent of process work;
* **handoff cost:** map seeding still allocates and performs ordered
  comparisons. The candidate removes parser replay, not the map's nodes or
  insertion work. Exactly-33 and 34-attribute rows must remain separate;
* **error-order drift:** a preflight that treats a missing equals sign as a
  duplicate, consumes the first `=` as a delimiter, or advances `next_start`
  after an error changes the public error contract; and
* **copy/API drift:** all five helper bodies and the shared OPC test oracle
  must remain synchronized, while OPC/OLE visibility and OOXML re-export
  surfaces remain unchanged. No call site should be moved to
  `unchecked_attributes()` as a shortcut.

The 0764 security record is a hard constraint here. A fixed direct prefix and
the existing ordered fallback are preferable to a new attacker-influenced
unkeyed hash. A candidate may be rejected for a large constant factor even if
its asymptotic bound is sound; the bound alone is not a performance result.
The 0764/0777 integration also keeps checked fail-fast readers distinct from
lenient `first_wins`, counting, and MCE-owned unchecked paths; this report does
not propose changing those caller classifications.

## Prior attempts and what they rule out

| record | attempted shape or evidence | retained consequence |
|---|---|---|
| [0777](../../0777-xml-attribute-integration.md) | integrated the bounded adapter across 545 fail-fast sites, with synchronized five-copy ownership | semantic/differential gates passed, but the isolated checked-iterator sublane used about 6.10–6.23% more instructions; OPC `n=30` was about 12.7–13.2% higher. Whole-process allocation rows were not parser attribution. This is an earlier warning that the ordered handoff and bookkeeping can dominate saved allocator traffic. |
| [0794](../../0794-xml-inline-duplicate-keys.md) | four inline borrowed names, `Vec` through 32, then ordered map; quick-xml checks disabled | allocation calls fell 85.048% (72,106 to 10,781) in the large PPTX capture, but no primary timing benefit; tiny XLSX full-cell scan regressed 5.511% and vetoed the candidate. Iterator size grew 120→192. Do not repeat this inline-four/vector design. |
| [0795](../../0795-xml-duplicate-key-instruction-profile.md) | instruction and branch attribution of 0794 | large-capture simulated guest instructions fell only 0.387–0.393%; checked-iterator/inspector self work rose and simulated indirect misses rose about 35%. It does not turn allocation reduction into a latency claim. |
| [0796](../../0796-xml-attribute-boundary-diagnostic.md) | explicit 0/1/4/5/8/9/16/17/32/33/64 and malformed tails | unchecked value-first duplicate remapping made 4,096-byte quoted/unterminated duplicates about 24.38–24.51× slower. Any new candidate must preflight before value scanning. |
| [0797](../../0797-xml-first-attribute-preflight.md) | unchecked first item, then replay to checked quick-xml when a tail remained | one-item consumption improved, but two-item and multi-item paths regressed; replay is not a viable general boundary. |
| [0798](../../0798-pptx-attribute-consumption-census.md) | production-shaped consumption census | 295,600 instances were fully consumed; large capture had 30,374 one-attribute and 30,816 two-attribute instances, 99.777% combined. This supports avoiding replay, but supplies no timing or allocation prediction. |
| [0799](../../results/change-0799/preflight-review.md) | no-replay first/two-item states and a third-item transition | direct one/two-item consumption improved 34.898%/16.716%, while six rows after the two-item boundary regressed about 1.50–1.56× and 19 rows were flagged. The experiment did not advance. |
| [0800](../../results/change-0800/source-review.md) | leading-`=` correction | fixed `name_at` so malformed late duplicates report `Duplicated(302, 2)` rather than `ExpectedQuote`; this is a correctness prerequisite, not an optimization result. |
| [0801](../../results/change-0801/source-review.md) | no replay for the first two, boxed ordered map after the third | short paths improved, but empty and 4/5/16/32 paths regressed, up to 68.034%; object size grew 120→128. Rejected. |
| [0802](../../results/change-0802/source-review.md) | boxed fixed array of 32 borrowed names, raw preflight, ordered handoff | semantics passed source review, but the boxed array still allocated and the candidate retained 19 direct-helper regressions. It is not an allocation-free answer. |
| [0803](../../results/change-0803/source-review.md) | moved exact-empty decision from construction to first request | isolated placement diagnostic only; it did not remove the array/map or parser costs. |
| [0804](../../results/change-0804/source-review.md) | direct byte equality in the bounded linear stage | equality versus `Name::cmp` was a useful diagnostic factor; it did not authorize workflow advancement or adoption. |
| [0805](../../results/change-0805/source-review.md) | combined no-replay, bounded array, direct equality, and lazy map | direct helper rows still regressed: distinct-4 +22.649%, duplicate-after-4 +10.726%, syntax-equals-after-4 +28.682%, syntax-flag-after-4 +28.634%, and distinct-32 +3.274%. It advanced only to workflow qualification. |
| [0806](../../0806-xml-workflow-qualification.md) | public five-copy workflow/resource/cross-format qualification | no capture or lifecycle benefit met the 3% gate; large lifecycle was 3.105% slower. Allocation calls fell 72,106→10,788, but production was restored. OLE visibility had to be repaired because the helper is public. |
| [0897](../change-0897/source-audit.md) | retained allocation attribution for the restored XLSX fallback | 85,560 duplicate-check stack events over 20 scans are a real path signal, but the report explicitly separates stack presence from removable allocation count and makes no latency claim. |

The inventory leaves one distinct hypothesis: retain only fixed raw key
positions and perform the early key preflight, rather than retaining values,
replaying values, or using an unkeyed hash. It is still a layout-sensitive
experiment, not an implied production change.

## Qualification gates if the experiment is resumed

No implementation is recommended from this source audit alone. If the root
agent elects to prototype the start-offset index, the first gate should be a
source and differential review against quick-xml, before any public workflow
run. The frozen matrix should include:

* zero, whitespace-only, one, two, three, four, five, eight, nine, 16, 17,
  31, 32, 33, 34, and 64 successful names, with adjacent and separated forms;
* duplicates at every boundary, including leading-`=` and non-ASCII raw keys;
* valid, unquoted, missing, quoted, and 4,096-byte unterminated duplicate
  values after one, two, 32, and 33 prior names;
* unique malformed tails, `ExpectedEq` cases, exact byte positions, borrowed
  values, clone continuation, repeated exhaustion, and `Debug` state; and
* direct counts for array initialization, `size_of::<CheckedAttributes>()`,
  construction, consumption, map seeding, allocation calls/bytes, net live
  bytes, peak bytes, and instructions. Diagnostic counters must not be
  interpreted as workflow latency.

The candidate must then pass the five-copy canonical-source and public quality
gates, the 0799 protected short/error rows, the 0806 capture/lifecycle and
cross-format policy, and the existing resource/hostile-input checks. A fresh
public trial would be required even if the direct helper shows an allocation
reduction. The 0798 count distribution cannot be multiplied by a microprobe
ratio to predict a workflow result.

## Disposition

The start-offset prefix index with pre-value raw-key checking is a concrete,
bounded, API-compatible mechanism that could remove the first quick-xml
duplicate-name allocation without storing inline attributes or adding a
dependency. Its costs are substantial: at least a 256-byte inline table on
64-bit targets, extra per-item bookkeeping, map-seed work at 33, and five-copy
source maintenance. The current preliminary native leaf share is below 1.1%
of the process, and the historical 0794–0806 record shows that large
allocation reductions did not produce an eligible workflow benefit.

Recommendation: treat this as a low-priority isolated experiment only if the
root allocation audit finds no higher-share removable path. Do not infer a
latency win from the 0897 call count, and do not carry any 0794/0802 inline
storage candidate into production without fresh short-tag, malformed-tail,
layout, resource, and public-workflow evidence.
