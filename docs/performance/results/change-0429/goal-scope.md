# Change 0429 goal scope

**Disposition: OPEN.** Change 0429 completes the 32-process provider and native-image
baseline, plus a bounded ZIP correctness fix. Formal reports and supplementary
profiles pass their recorded checks; portable replay supplies the final bundle gate. Nothing in this batch authorizes a
speedup, regression, memory-release, or global-goal claim.

## What this batch covers

| Area | Frozen scope | Evidence boundary |
| --- | --- | --- |
| Synthetic source providers | `plain` and `media-rich`, each through `bytes`, recently written/reopened positional `file`, 4,096-byte capped `range`, and 65,536-byte capped `range` with a configured 200 µs delay | The file lanes are warm positional-file observations. The range delay includes scheduler overhead and is an instrumented caller condition, not a remote or calibrated bandwidth measurement. |
| Native selected-image inputs | The two unmodified POI fixtures, `poi-slide` and `poi-video`, through the same four provider forms, with retained image payloads and independent ZIP/XML oracles | This is a selected-image lifecycle and preservation probe. It does not establish native producer breadth or a positive native cross-copy workflow; the inspected unnamed native slides remain a typed refusal boundary. |
| Formal sampling | 32 fresh release processes, two repeats, three warmups and 30 retained samples per process: 960 retained samples; supplementary 100-iteration media profiles are outside that count | The protocol and command receipts must establish completion. The count cannot be inferred from the frozen protocol alone. |
| ZIP short-read correction | `crates/soapberry-zip/src/archive.rs` refills only missing fixed-header bytes, parses already buffered headers, and rejects a truncated trailing fixed header. Development receipts record four generated regressions passing after the fix and six expanded capped-read tests passing | This enables valid short-read providers and closes a correctness defect. It supplies no latency or I/O-efficiency result. |

The clocks cover the named public API sequences only. Source construction,
copying, sink reservation, checks, observers, corpus gates, and drops are
outside those clocks; process RSS includes setup and observers. Logical
`ReadAt` request histograms, managed budget/cache observations, and RSS remain
separate evidence streams. These boundaries are recorded in the
[0429 protocol](protocol.json) and [source review](source-review.md).

## Acceptance status before 0429 can be retained as evidence

The two native source/verifier issues identified during review have been
fixed: the first available read interval now has a checked zero-baseline delta,
and the selected-image descriptor oracle checks the promised identity fields.
Focused native harness tests pass for both POI fixtures and their three tested
provider forms; the shapes fixtures retain their refusal controls. Final
source review and captured-report verification pass. Native evidence remains
limited to these selected-image lifecycles.

The capture must retain raw reports and catalogs, exact corpus and fixture
hashes, provider configuration, output/payload preservation gates, source-read
histograms, and per-phase ownership observations. A passing formal capture
must be established from the retained verification and replay receipts, not
from the planned sample count. See [native review](native-review.md) and the
[development corrections](checks/development-corrections.md) for the current
correction history and refusal controls.

## Remaining priorities

1. Improve CPU stack completeness and setup/check attribution before selecting
   a new optimization from the completed provider baseline. Retain separate
   API-duration, logical-read, managed-ownership and RSS conclusions; do not
   collapse the API timer into an end-to-end lifecycle claim.
2. Use the 0429 source boundary to design the next cold/range/native extension:
   separately prove filesystem cache state or a controlled caller source, and
   retain the same selected-image/output oracle. The current recently written
   file lanes and deterministic delay do not answer that question.
3. After that evidence is clean, extend the highest-impact missing non-iWork
   workflows with comparable end-to-end boundaries and attribution: broader
   native producer/size coverage, representative CRUD gaps, bounded semantic
   streaming/append, and then explicit CPU/scaling measurements. The existing
   representative CRUD, strict-gate, and broader streaming/scaling gaps remain
   open as described by the [global audit](../../../GOAL_AUDIT.md).

The 0429 ZIP change is therefore a correctness enabler for the next provider
measurements. It does not by itself close cold I/O, native producer, allocator,
physical-copy, general retention, scaling, or the comprehensive non-iWork
CRUD goal.

The supplementary profiles resolve only a minority of iteration ancestry.
Improve stack completeness and retain setup/check attribution before ranking
a new optimization from that CPU evidence; see [profile review](profile-review.md).
