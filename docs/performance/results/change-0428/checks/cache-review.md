# 0428 managed PPTX cache and budget review

This is an independent, source-only review of the managed PPTX cache-retention
design and the current `cache-retention` harness. I read the 0428 protocol and
design note, the 0086/0088 cache records, ADRs 0003, 0005, 0006, 0008, and
0024, and the source-backed PPTX/OPC owners. I did not run Cargo, tests,
builds, profiling, or capture commands.

## Result

The lifecycle ownership model is sound in the current source. The source
presentation and destination editor receive separate caller-owned budgets and
execution contexts. A cross-copy plan retains a source presentation clone and
its staging reservation; publication consumes the destination editor. The
current lifecycle journal samples the destination editor before that consuming
call, records the surviving source view and caller budgets after it, and drops
the returned result, plan, view, caller sources, and sink in a useful order.
It uses the fallible diagnostics API and represents consumed or dropped owners
with unavailable values instead of fabricated zeroes.

The earlier snapshot had a release blocker: `run_capture` rejected every
scenario except `lifecycle`. The frozen harness now dispatches the five near
scenarios through `run_near_capture`, builds the selected corpus before the
warmup/sample loop, and emits the scenario-specific journals. That initial
finding is retained below as resolved history. The Rust correction serializes
actual A/B-pair and B-reload read deltas for the pinned row. Direct image rows
carry `Some` payload counters; publication-only repeated rows carry explicit
`None` for the non-applicable image and reload counters. The independent
`verify-report.py` consumer now validates that optional-field contract,
compares every present read oracle with its phase interval, and rejects
fabricated reload values. It also enforces the prescribed cache-owner and
source-read availability at every phase and matches each available cache
diagnostic's used resources to its caller budget. No source or verifier
blocker remains in this review.

## Ownership and budget checks

The following source facts agree with ADRs 0003 and 0005 and with the 0086 and
0088 managed-cache contracts:

* `SourceBackedPresentation` and its slides share one source package through
  `Arc`; a slide or image handle is therefore an intentional owner that can
  keep cache state alive. The source-backed cross-copy plan also owns a source
  presentation clone. A descriptor or metadata value alone must not be treated
  as a payload pin.
* `SourceBackedPresentationEditor::publish_cross_slide_copy_to_stream` consumes
  the editor. Its last valid destination-cache snapshot is immediately before
  publication. The destination cache must remain unavailable after that point;
  the caller-owned destination budget is the valid post-consumption observation
  of current reservations.
* The current lifecycle code drops the returned publication result before the
  plan and drops the plan before the source view. That matters because the plan
  retains the source owner and its destination staging reservation. The source
  and destination `Arc<InstrumentedSource>` values are moved into the two
  package constructors and retained only as caller observation handles after
  construction, so they do not create a second managed package.
* `validate_budget_gauges` checks each owner’s diagnostic resource gauges and
  limits against its own `Budget`, and bounds cache reservations and retained
  entries by the configured cache limits. Source and destination values remain
  separate. `Memory` and `Objects` are current reservations; `InputBytes`,
  `OutputBytes`, and `Work` are cumulative charges and must not be expected to
  return to zero after a drop.
* `try_cache_diagnostics` and checked counter deltas are the correct evidence
  APIs. A poisoned owner, diagnostic overflow, or owner consumed by publication
  must invalidate or omit that owner’s point. The current unavailable points
  preserve this distinction.

The corrected design choices are appropriate: independent source and
destination roots, a 64 MiB/2-entry pinning row that accounts for the pinned
`_presentation` root, and a `P - 1`/128 oversized-bypass row that is distinct
from a one-byte-under managed-memory refusal. The parser also currently allows
`repeated-publication` with the plain corpus (`requires_media_rich` at
`pptx_cache_retention.rs:60-62`), matching the protocol.

## Initial findings and current disposition

### Initial blocker: near-image scenarios had no execution path (resolved)

The initial code accepted the five near names in `Scenario` but rejected them
at capture. The frozen code now has checked dispatch for exact admission,
one-under, pinned eviction, oversized bypass, and repeated publication. It
records their actual per-row cache and budget limits, while the unit gate still
rejects invalid image/plain combinations and accepts repeated/plain.

### High: refusal rows need a target-payload read oracle (implemented; retain the boundary)

The current `ReadPoint` is a total counter over an `InstrumentedSource`; it
does not identify a byte range. That is sufficient only if the formal baseline
is taken immediately after the selected slide’s metadata (`images`/`image`)
and no other source operation occurs before the target request. The one-under
row must compare calls and bytes against that baseline and prove a typed
`Resource::Memory` refusal, no target cache entry/reservation, and zero sink
bytes. If the implementation performs any metadata lookup during
`read_image`, the baseline must be after that lookup, as the design note
requires. Do not label a total-counter delta as a target-range proof unless
the phase ordering makes that implication exact.

The frozen image path takes the `ReadPoint` immediately after validated
metadata, then checks the metadata-to-payload delta. Its typed error matcher is
restricted to `Resource::Memory`, and the one-under branch requires zero read
delta, zero output, and the root-only current memory gauge. This is adequate
for the fixed synthetic corpus; it remains a logical positional-read oracle,
not a physical-storage or decompression oracle.

The reservation retry path makes the revised calibration precise. At
`crates/litchi-opc/src/source_backed.rs:2885-2916`, a failed memory reservation
evicts all unpinned clean entries and retries before returning the error. The
PPTX source facade retains the mandatory presentation root in
`SourceInner::_presentation`; the selected slide metadata is loaded only for
the query and its `SourcePart` is dropped before `read_image` reserves the
image payload. Therefore the formal memory cap should be derived as:

```text
root_floor = source budget Memory immediately after the source view opens
exact_cap  = root_floor + declared selected payload P
under_cap  = exact_cap - 1
```

This is a `Resource::Memory` reservation boundary, not a process-memory or
allocator measurement. The root floor remains after `evict_all` because the
root handle is externally pinned; the metadata entry is clean and evictable.
At the exact cap, the first reservation may fail while metadata is resident,
then succeed after metadata eviction. At the one-under cap, the retry still
cannot fit alongside the pinned root, so the request must fail before the
target payload `ReadAt`. The evidence should accept the actual exact-row
failure/eviction counter rather than requiring a fixed first-attempt count,
while requiring the one-under row's typed memory refusal, zero target-read
delta, and no target cache/reservation. A public metadata-payload pin is not
needed and would invalidate the intended eviction/retry boundary.

### High: near rows must make handle ownership explicit (resolved)

The pinning and bypass rows need named `SourceSlide`/`SourceImage` scopes and
must drop those handles before interpreting cache and budget release. With the
root pinned and image A pinned under `SourceCacheLimits::new(64 MiB, 2)`, a
request for B can bypass both selected-slide metadata and B’s payload; the
event count is therefore not necessarily one. A returned bypass handle can
consume `budget_memory_used` while contributing nothing to
`budget_cache_reserved_bytes`. The oversized row likewise must show successful
uncached data, an oversized-bypass event, and a cold second read without
calling it a budget refusal. The clean-eviction row must verify payload bytes
and read transitions, rather than infer correctness from an eviction counter.

The frozen pin implementation precomputes A once from the corpus builder,
uses the corpus selected position and selected payload for B, keeps both image
handles alive through the pinned-B request, drops both before the clean reload,
and checks the root-plus-A/B budget and cache gauge transitions. Its row-level
payload fields now contain the actual A/B-pair delta and its reload fields the
actual B-reload delta. The oversized row retains and drops each uncached image,
then proves a cold reload under the `P - 1`/128 policy. The repeated row
explicitly leaves both image/reload counter pairs `None`, since it has no image
payload phase. The verifier now requires `Some` payload counters for image
scenarios, requires reload counters only for oversized and pinned rows, requires
`None` for repeated publication, and compares every present value with its
phase delta. It also rejects unavailable-owner points that contain fabricated
cache or read counters and checks available cache gauges against their caller
budget.

### Medium: completed-owner invariants are captured and enforced by the verifier

The Rust capture records these gauges and the independent report verifier now
rejects nonzero `in_flight_loads` at completed points and nonzero final
`Memory`/`Objects` after all named handles are dropped. It continues to check
`InputBytes`, `OutputBytes`, and `Work` as cumulative monotonic charges rather
than requiring them to return to zero.

### Medium: repeated publication must separate plan reservations from output (satisfied)

The repeated row must reuse one source view and source context, create a fresh
destination editor each iteration, and retain the external destination budget
after each consuming publication. The plan’s staging reservation can survive
editor consumption until `drop(plan)`, so destination `Memory`/`Objects` must
be checked only after the documented result/plan/sink drops. The three-success
`OutputBytes` ceiling is set to three times the exact accepted publication
length and reported as a cumulative charge; it is not a retained-memory or
exact-peak-memory claim. The fourth operation is a separate typed
`OutputBytes` refusal with a zero-byte sink, and the destination releasable
budget is checked after its plan is dropped.

## Frozen near-case checks

The root-floor calibration remains sound after the near implementation landed.
`calibrate_image_memory` opens the same source archive and cache policy,
records `Memory` immediately after the presentation view opens, then runs the
same metadata oracle before dropping the calibration owner. Formal exact and
one-under rows use `root_floor + P` and `root_floor + P - 1`, respectively.
The production cache may evict clean metadata between its first and retry
reservations, so exact admission is allowed to show an initial reservation
failure/eviction while one-under must fail on the retry before target I/O. No
metadata payload pin is introduced.

The repeated row keeps one source view and its source budget alive across three
fresh destination editors. Each plan/result/sink scope is dropped before the
next editor, while source diagnostics continue from the same owner. The fourth
refusal is attempted at the exact cumulative output ceiling. The near capture
does not install a callback observer or leak observer storage; diagnostics use
the public fallible snapshot seam, and corpus construction occurs before the
warmup/sample loop.

## Acceptance boundary

Before formal capture, retain the source revision, binary identity, corpus
gates, actual cache and budget limits, source read counters, typed
success/refusal result, exact output bytes/hash where applicable, and explicit
ownership/drop labels. Destination cache diagnostics after publication must
remain unavailable, while source diagnostics are sampled only while a source
owner remains. The independent verifier now accepts only the Rust row schema's
applicable optional image/reload counters: image rows carry `Some` deltas,
repeated publication carries explicit `None`, and phase checks validate each
present interval. It also requires live cache/read availability and exact
caller-budget gauge matching at available points. RSS/VmHWM points, if
retained, are process observations with probe overhead and cannot support a
cache-release or performance claim.

No finding here changes the production ownership model or the corrected 0428
design. The initial harness-completeness blocker, Rust pinned-row counter
defect, and verifier optional-field handling are resolved. This source-only
review has no remaining known blocker. No Cargo/build/test result is claimed
by this review.
