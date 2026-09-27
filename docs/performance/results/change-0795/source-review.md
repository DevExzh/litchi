# 0795 source review — instruction, branch, and layout paths

This review defines the source-backed diagnostic questions for the 0795
Callgrind and native sampled-stack replay. It does not change production
source, apply the archived candidate, run a build, or claim a latency cause.
The current worktree is the exact 0794 baseline at `65fdca1a04`; production
must be restored to that baseline before either leg is captured. The candidate
leg is the rejected 0794 archive under
`change-0794/candidate/after/`, bound by the 0795 inheritance record.

The five candidate files are the fail-fast checked-attribute copies in
`litchi-opc`, `litchi-ole-common`, `litchi-sign`, `litchi-xldm`, and
`xml-minifier`. The formula module's lenient `first_wins` helper is a separate
contract and is not part of this diagnostic. The 0794 decision remains in
force: its candidate saved operation-scoped allocations but had no eligible
primary timing benefit and triggered the small XLSX full-cell-scan veto.

## Evidence and constraints

The baseline wrapper in `crates/litchi-opc/src/xml_attributes.rs` delegates
the first 32 successful attributes to quick-xml's checked `Attributes`. At
the 33rd request, `next_after_quick_xml` disables quick-xml checks, reparses
the first 32 names with an unchecked iterator, constructs the ordered map, and
then checks the current item. The candidate archive instead disables checks
when `CheckedAttributes::new` constructs the iterator and performs every
duplicate check in `OwnCheck`.

The candidate's `SeenNames` states are source-visible and bounded:

* `Inline` stores four borrowed `Name` values in an array;
* the fifth distinct name allocates a vector with capacity eight, which stays
  linear through 32 names; and
* the 33rd distinct name moves that vector into the existing `BTreeMap`.

The 0794 count-only evidence gives a useful boundary map without being an
instruction result. For 1–4 attributes, the baseline recorded one 64-byte
allocation and the candidate recorded none. At 5–8, 9–16, and 17–32, the
baseline/candidate allocation-call pairs were respectively 2/1, 3/2, and
4/3; at 33 they were 10/8, and at 64 they were 14/12. The candidate iterator
size was 192 bytes versus 120 bytes for the baseline. These are separate
resource observations. They do not establish that a saved allocation is the
source of any observed timing difference.

The public notes inspector returns before constructing the checked iterator
when `element.attributes_raw().is_empty()` (the retained 0792 empty-tail
branch). A public capture therefore contains both an unaffected empty-tail
path and a checked-attribute path. The 0795 owner region must not be described
as if all capture work passed through this helper.

The relevant accepted constraints are the measured-evidence requirement in
ADR 0005, the malformed-input and fail-closed requirements in ADR 0006, and
the dependency/ownership rules in ADRs 0002, 0010, and 0011. The diagnostic
must preserve the same corpus, owner boundary, source identity, and leg order
as the frozen 0795 plan. It cannot be used to retain the rejected candidate or
to weaken the public performance gate.

## Exact paths to compare

The following cases are the smallest source-backed decomposition of the
candidate. They should be inspected in the owner-scoped Callgrind output and,
where the public workload reaches them, in the native stacks. If an auxiliary
direct-tag diagnostic is added later, it should be a separately labelled
artifact; it must not silently change the frozen tiny/medium/large public
matrix.

| Case | Baseline path | Candidate path | Counters or stacks to inspect |
|---|---|---|---|
| Empty attribute tail | `CheckedAttributes::new` is not reached by the notes inspector; the caller returns after namespace/root checks. | Same caller branch; no `SeenNames` state is consumed. | Confirm that this branch does not acquire helper frames. Any public delta here belongs to surrounding capture work. |
| Iterator construction and zero attributes | `CheckedAttributes::new` calls `tag.attributes()` and creates `Phase::QuickXml(0)`. The first `Attributes::next` reaches end of input without key storage. | `CheckedAttributes::new` calls `unchecked_attributes()` and `OwnCheck::new`; `SeenNames::new` initializes the inline array and `next_start` from `tag.name()`. The first `next` matches `Phase::Own`, sees `None`, and enters `Done`. | Constructor `Ir`, prologue/return stores, and the `None`/`Done` branches. This is the direct layout-cost control; it is not an allocation-saving case. |
| One through four distinct names | quick-xml's `IterState::next` runs `check_for_duplicates`, linearly scans `keys`, and pushes a `Range`. The wrapper only increments its `QuickXml` counter. | `IterState::next` parses with checks disabled; `OwnCheck::check` calculates `position`, calls `SeenNames::check` in `Inline`, compares raw bytes, and calls `end_of` after acceptance. | `IterState::check_for_duplicates` versus `SeenNames::check`, `offset_in`, `end_of`, and the slice-equality/length path. The candidate computes an offset and advances `next_start` for every accepted small attribute; the baseline wrapper does not do those wrapper operations in its first phase. |
| Fifth distinct name | The quick-xml key vector continues its normal capacity/growth path. | `SeenNames::Inline` copies four `Name` entries into a vector with capacity eight and pushes the fifth. | Allocation/growth frames, the copy/initialization path, the `len < INLINE_NAMES` transition, and branch counts. The 0794 allocation result identifies the boundary but does not say whether copy, branch, or allocator work dominates. |
| 8→9 and 16→17 | quick-xml's key vector grows according to its own `Vec<Range<usize>>` capacity history. | The candidate vector grows from its explicit capacity-eight spill at 9 and again at 17. | `RawVec` growth, element copy, `SeenNames::Linear` scan, and the capacity comparison. These exact thresholds are not isolated by the public shape names, so a public zero count must not be interpreted as proof that they are absent from the package. |
| 31, 32, and 33 distinct names | The first 32 items stay in quick-xml. The 33rd request enters the cold, non-inlined takeover, reparses 32 names, allocates the ordered map, boxes `OwnCheck`, and checks the 33rd item. | The first 32 items remain in the linear vector. The 33rd check takes the vector with `mem::take`, inserts the retained names into the ordered map, inserts the current name, and keeps `OwnCheck` inline in `Phase`. | `next_after_quick_xml`, `take_over`, `OwnCheck::after_quick_xml`, repeated unchecked parsing, and `BTreeMap` insertion versus `SeenNames::check`'s transition. This is the strongest direct comparison of reparse work and layout/boxing, but it may be rare in the public corpus. |
| More than 33 distinct names | Baseline `OwnCheck::check` uses `BTreeMap<Name, usize>::entry` after takeover; quick-xml duplicate checking is already off. | Candidate `SeenNames::Ordered` uses the same ordered map and `entry`; its `Name::Borrow<[u8]>` lookup is used by malformed-error recovery. | Ordered-tree comparisons, node allocation, `Name::cmp`, `Borrow` lookup, and the map-entry branches. Both costs are nested under the owner and must not be summed with parent counts. |
| Early valid duplicate | quick-xml checks the name before parsing its value, returns `Duplicated`, and records `SkipEqValue` recovery state. The wrapper stops at the first error. | The unchecked parser first returns the complete lexical item; only then does `SeenNames::check` return `Duplicated`. | Compare the parser's value scan and the local duplicate branch. A long quoted or unterminated value can be scanned by the candidate before the same typed error is returned. This is a source-level work difference, not a measured public-workflow cause until counters isolate it. |
| Duplicate with malformed value | Baseline reports the duplicate before value parsing. | The candidate receives an unchecked `AttrError`, then `duplicate_before_value` calls `name_at` from the last accepted value and `SeenNames::position`. | `duplicate_before_value`, `name_at`, `position`, and the `UnquotedValue`/`ExpectedValue`/`ExpectedQuote` match. Check the inline, spill, and ordered stores separately; a short-value test cannot establish equal work for a long tail. |
| Non-duplicate syntax error | quick-xml emits `ExpectedEq`, `ExpectedValue`, `UnquotedValue`, or `ExpectedQuote` from its parser and the wrapper enters `Done`. | The unchecked parser emits the same lexical error; only the value-related variants enter the candidate's duplicate-recovery search. | Error-variant branch paths and `Done` transition. `ExpectedEq` must remain distinct because duplicate checking has not reached a name followed by `=`. |
| Post-error, exhaustion, and clone | `Phase::Done` is reached after the first error or `None`; `Clone` copies the quick-xml state and whichever phase is live. | The same public behavior is implemented with an inline `OwnCheck`, including array/vector/map clone and drop paths. | These paths are unlikely in the timed capture, but they are the layout-sensitive operations to inspect in any direct diagnostic. No public clone cost should be inferred from the capture profile. |

The most important extra instruction paths are therefore `end_of` and
`offset_in` on every accepted candidate attribute, the raw-byte scan in the
inline/linear stores, and the candidate's complete lexical value scan before a
valid duplicate is recognized. The most important branch paths are the
candidate's `Phase::Own`/`Done` match, `SeenNames` state dispatch, inline
length and vector-capacity transitions, and the error-recovery match. The
layout path is the 72-byte larger `CheckedAttributes` value and its inline
`OwnCheck`; it affects construction, moves, clones, and the stack/ABI even
when the tag has no attributes.

These are candidate costs and possible explanations to test. The source does
not establish which one accounts for the 0794 public p50 result.

## Callgrind comparison semantics

The frozen 0795 Callgrind packet uses one positive owner-scoped dump per
shape, the tiny/medium/large order in both directions, and the before/after
leg orders in both directions. The stable owner is the non-inlined
`namespace_uri_probe::capture_region_0793` wrapper. Every report should bind
the source and binary identities to its leg and retain the exact command and
termination receipt.

Interpret the events as follows:

* `Ir` is Valgrind's simulated guest instruction count in the collected owner
  region. It is not elapsed time, retired hardware instructions, or CPU cycles.
* `Bc`/`Bcm` are Callgrind's conditional-branch and simulated conditional
  misprediction counts; `Bi`/`Bim` are the corresponding indirect-branch
  counts. They are useful for comparing this same binary/protocol pair, but
  they are not hardware branch-predictor measurements.
* A delta is meaningful only for the same shape, owner scope, event, source
  identity, build settings, and collection boundary. The order reversals
  control for process/order drift; they do not create a confidence interval
  from one dump per cell.
* Nested parser and allocator counts are included in the owner region. A
  parent `CheckedAttributes::next` count and a child `IterState::next` count
  overlap. Report them as separate annotated rows and never sum them into a
  removable fraction.

The public tiny/medium/large shapes test the end-to-end instruction mix. They
do not prove how many attributes fall at 4, 5, 9, 17, or 33, and a zero
Callgrind line for a transition is not evidence that the source path cannot
execute. If the driver has an isolated tag probe, use exact controls at
`0, 1, 4, 5, 8, 9, 16, 17, 32, 33, 64` attributes plus valid and malformed
duplicates. Keep those reports separate from the public owner capture.

Useful line/function groups for the before leg are:

* `CheckedAttributes::new` and `CheckedAttributes::next`;
* `next_after_quick_xml`, `take_over`, and `OwnCheck::after_quick_xml` (the
  takeover is explicitly cold and non-inlined); and
* quick-xml `Attributes::next`, `IterState::next`, and
  `IterState::check_for_duplicates`.

Useful groups for the candidate leg are:

* `CheckedAttributes::new`, `OwnCheck::new`, `CheckedAttributes::next`, and
  `OwnCheck::check`;
* `SeenNames::check` and `SeenNames::position`, including the `Inline`,
  `Linear`, and `Ordered` source lines;
* `offset_in`, `end_of`, `name_at`, and the ordered `BTreeMap` entry/compare
  code; and
* quick-xml `Attributes::next`/`IterState::next` with duplicate checks turned
  off.

The candidate should not logically enter quick-xml's duplicate-check function
because `with_checks(false)` is set at construction. Generic parser code or
symbol tables may still retain a branch or an inlined body, so absence or
presence of a symbol must be checked against annotated source and disassembly
metadata rather than treated as an execution count.

## Native sampled-stack semantics and expected visibility

The 0795 native lane is a separate large-capture run at `cycles:u`, 499 Hz,
three warmups, and 100 samples per leg/order. It is frame-pointer sampled
work under the same `capture_region_0793` owner. It is not the Callgrind
instruction counter and it is not a native latency measurement.

The baseline has a useful stable observability advantage: the takeover helper
is marked `#[cold]` and `#[inline(never)]`, so `next_after_quick_xml` and
`OwnCheck::after_quick_xml` may appear as distinct frames when the large
workload reaches the 33-name path. The normal first-32 path may instead show
only `CheckedAttributes::next` and quick-xml's iterator frames.

The candidate's `SeenNames::check`, `position`, `offset_in`, and `end_of` are
small ordinary helpers without an `inline(never)` anchor. Thin LTO and release
optimization may fold them into `CheckedAttributes::next` or `OwnCheck::check`;
the candidate can therefore perform the work without exposing a separately
named frame. Conversely, map comparisons may appear as monomorphized
`BTreeMap`/`Name::cmp` or allocator frames rather than a source-level
`SeenNames::Ordered` symbol. The candidate's absent `next_after_quick_xml` is
expected and does not itself demonstrate less work.

The following observations are safe to make if the stacks support them:

* quick-xml duplicate-check frames are expected on the baseline's first-32
  path and should be absent from the candidate's actual duplicate-check path;
* candidate `SeenNames` or its inlined source lines are expected under the
  owner for nonempty tags, while baseline `OwnCheck::after_quick_xml` is
  expected only at the takeover boundary; and
* allocator/growth frames and ordered-tree frames should be interpreted only
  for stacks that also carry the owner anchor and only as observed samples.

A zero sample for a helper does not prove zero execution. An unresolved or
inlined frame does not support attributing its work to a named source helper.
The packet should retain unresolved-frame counts, symbolizer/decode settings,
and the exact owner filter. A phase fraction or owner share is admissible only
if every relevant owner frame in the selected stacks is resolved; otherwise
report raw observed stack rows and leave the fraction uncomputed. Stack counts
overlap by construction and must not be added across ancestor/descendant
rows.

Only the `litchi-opc` copy is expected to be exercised by the PPTX namespace
probe's normal OOXML path. The OLE common, signature, XLDM, and XML-minifier
copies remain source-synchronized and quality-checked, but this owner-scoped
PPTX sample cannot establish their runtime path frequencies. The formula
helper must not be included in either leg.

## Layout and semantic comparison

The known 120→192 byte iterator-size change is stronger layout evidence than
a stack prologue impression. A future direct layout row should use the same
compiler/target and report `size_of::<CheckedAttributes>()` for each leg. It
should keep this separate from Callgrind `Ir` and native samples: a larger
Rust value can alter register/stack moves without producing a stable function
symbol, and a smaller constructor does not imply a smaller live working set.

The semantic compare remains exact checked-iterator output through the first
error. In particular:

* raw QName bytes remain the duplicate key and positions remain byte offsets;
* the first error and every post-error `None` remain the same;
* `ExpectedEq` stays a syntax error, while malformed duplicate values retain
  duplicate precedence through `duplicate_before_value`; and
* the candidate's long quoted/unterminated duplicate scan is a work-order
  difference even when the resulting `AttrError` matches the baseline.

Callgrind and perf cannot replace the differential and package quality gates.
If a diagnostic row shows fewer quick-xml instructions but altered errors,
limits, namespace handling, or borrowed-value behavior, it is not an
optimization result and the source must remain rejected.

## Reporting boundary

The 0795 report may say that an owner-scoped diagnostic observed a change in
simulated instructions, simulated branches, branch-misprediction estimates,
or sampled stack counts for a named shape and leg. It may say that the source
contains the candidate's inline, spill, map, value-scan, and larger-layout
paths described above.

It must not say that any one counter caused the 0794 p50 regression, that a
sampled stack is an instruction fraction, that Callgrind counts are hardware
cycles, that an absent frame proves no execution, or that the diagnostic
justifies production adoption. The public rejection and exact baseline
restoration remain the controlling decision.
