# Change 0494 measurement review

Review date: 2026-09-10. This is an independent read-only review of the
0494 provider harness, `measure.py`, `cold_measure.py`, `profile.py`, and the
DOCX provider source. It owns this review file only. No build, capture, or
source edit was performed for this review, and the protected `litchi-spec-gaps`
worktree was not inspected or modified.

## Review result

The source-bound formal and pilot evidence is complete, and the acceptance
verifier passes the warm and cold lanes, retained-data recovery, the corrected
observer receipt, final gates, and current helper custody. Warm and cold
`pilot2` each have a frozen protocol, source/build bindings, complete pilot
inventories, and passing analysis and verification receipts. The warm formal
lane has 24 source-bound terminals and 720 measured rows; the cold formal lane
has four source-bound cells and 120 eligible invocations, all with passing
analysis and verification receipts. The accepted-evidence manifest assembles
those lanes and the verifier independently audits strict cold allocator
conservation. It selects the sole passing retained-data recovery attempt and
the highest passing observer correction rather than trusting a fixed attempt
number.

The selected observer correction is
`profiling-counter-recovery-r5/observer-correction.json`, SHA-256
`da1b57f2f63646ff081f979b684091c4dec05100c66604d6957a09c803843fac`. Its
post-processing gate exits 0 with unchanged source custody; the final7 helper
gate and helper inventory also pass. At the time of this review, a read-only
`verify_bundle.py verify` reached only the cleanup proof and stopped because
the retained build cache `/home/zhuhe/.cache/litchi-build-0494` still existed.
That is the remaining operational acceptance step; no measurement or observer
custody blocker remained. There was no accepted seal in the checkout at this
review point. Earlier smoke and `pilot1` development receipts remain
historical because they were produced while the source or drivers were
changing.

I independently re-collected the warm formal terminals and recomputed every
raw vector, percentile, and bootstrap receipt: all 24 cells and 720 rows
matched the retained analysis with zero differences, and all 360 allocator rows
satisfied the live-byte and peak equations. The formal repeat-variance receipt
must remain visible: normal `instrumented` p50 elapsed time moved from
2,298,452 ns to 4,521,263 ns (+96.7%), normal `file-warm` moved from
4,721,174 ns to 2,317,535 ns (-50.9%), allocator `instrumented` moved from
5,025,646 ns to 8,226,288 ns (+63.7%), and allocator `file-warm` moved from
5,280,968 ns to 7,230,353 ns (+36.9%). Read counts and output bytes stayed
stable, so these are repeat-variance observations rather than optimization
evidence. I independently re-collected the four cold formal cells and
recomputed their cell vectors, percentiles, bootstrap receipts, and inventory:
all values matched with zero differences, and all 60 allocator rows satisfied
the live-byte and peak equations. Cold normal elapsed p50 moved from
148,065,004 ns to 238,429,900 ns (+61.0%) between repeats; cold allocator p50
moved from 98,139,070 ns to 93,802,031 ns (-4.4%), with its p95 moving
-29.1%. Cold source calls (379), process read bytes (16,793,600), and output
bytes (16,793,612) remained stable. These are same-cold-cell repeat
observations only.

The intended warm matrix is six explicit provider arms (`owned`,
`instrumented`, `file-warm`, `short`, `delayed`, and `range-zero`) across normal
and allocator binaries, two repeats, and 30 retained rows: 24 child processes
and 720 measured rows. The pilot is 12 processes and 36 rows. The separate cold
lane is four role/repeat cells with 30 fresh parent invocations and one fresh
measured child per invocation, for 120 eligible rows; its pilot is two cells
and six invocations. A cell is eligible only after its own source-stable build,
protocol, custody, raw report, analysis, and verification receipts pass.

The four formal and pilot verification receipts also re-hash cleanly: every
warm report/terminal pair and every cold aggregate report matches its retained
receipt hash. No terminal or aggregate inventory drift was found.

The acceptance verifier's cold audit re-collects the retained formal and pilot
raw reports, checks strict unsigned-integer fields, live-byte conservation,
peak ordering and bounds, reallocation-call bounds, and exact allocator
inventory (60 formal plus 3 pilot rows). The canonical `cold_measure` validator
is weaker when called directly, so that API remains a defense-in-depth issue;
the acceptance authority must continue to run this strict raw-report audit.

## Timing and allocation boundaries

The warm Rust clock begins after corpus construction, provider construction,
source-version setup, sink reservation, and the allocation region setup. It
covers source-backed DOCX open, one paragraph replacement, commit, sequential
publication, commit diagnostics, source/candidate XML identity comparisons, and
the explicit commit drop. Output hashing, semantic/media verification,
preflight patch oracles, counter snapshots, and provider teardown are outside
the elapsed vector. The allocator region starts before the clock and is
finished after it, so allocator rows include an absolute process baseline and
post-clock finish work. `region_peak_live_bytes - live_bytes_before` is the
meaningful peak increment; absolute live/peak values are not per-operation RSS.

The cold child opens `FileSource` inside the timed lifecycle and keeps cold
preparation, expected-value construction, sink/counter reservation, output
verification, and patch oracles outside. Its process `/proc` bracket is child
local; the after probe includes procfs overhead. GNU time is a parent-plus-
waited-child receipt and must not be described as operation-local RSS.

The compatibility constructor is deliberately unmanaged. The report should
retain the actual `ReadLimits::default()` and `SourceCacheLimits::default()`
values while marking `resource_budget.managed=false` and all managed budget
fields unavailable. The current source does this; no managed memory, input,
output, object, depth, or work claim is admissible.

The zero-length scope has two layers in the implementation: `CountingReadAt`
delegates an empty call to its immediate wrapped provider, while
`PptxRangeSource::read_at` returns `Ok(0)` before invoking its physical source
or adapter counters. The range arms therefore count an outer logical empty
call without a corresponding physical call. The README and methods record this
distinction, so no physical-empty equivalence is inferred.

## Provider and counter contract

`owned` intentionally has no logical counter. `instrumented` and `file-warm`
retain caller-visible logical ranges. `short` adds a 4,096-byte cap and both
logical and physical traces. `delayed` uses a 65,536-byte cap, a 1,000 µs
fixed delay, 104,857,600 bytes/s transfer pacing, and minimum-service timing.
`range-zero` uses the same range and rate controls with a zero-duration fixed
delay, retaining it as a range-control arm rather than calling it an
unmeasured zero-latency source. `file-warm` is explicitly a recently-written,
warm-cache observation; it carries no cold-filesystem claim.

The Rust source retains complete logical and physical `ReadCounter.ranges` in
warm and cold reports. The analyzer can therefore recompute request-size
histograms and media overlaps from raw ranges. Accepted analysis must preserve
that distinction: logical caller requests, physical adapter requests, and
`PptxRangeSource` service counters are different observations. Physical
`read_bytes` in the cold lane is the admission signal, not a physical-media
claim.

## Findings that block acceptance

The 0484 CPU lock path is intentional shared historical-lane serialization;
the bundle's development notes document that choice. Likewise,
`ReadStats.snapshot` rejects any trace whose retained range count differs from
the call count, so an overflowing range cannot produce a later valid snapshot.
Those are review notes rather than blockers.

The post-build warm validator now reconciles each retained logical/physical
range pair with the configured cap, fixed-delay count, transfer-paced count,
nominal transfer delay, and all aggregate byte/call totals. It also validates
strict boolean fields, the exact patch-oracle scope, allocator live/peak
conservation, and physical media-overlap vectors. Cache analysis now retains
only fields emitted by Rust, including the boolean materialization oracle;
absent cache measurements are no longer synthesized as zeroes. Pilot2 exercised
these checks across all six arms and both roles, and its 36 measured rows were
accepted by the canonical validator.

1. Profile report identity delegates to canonical `measure.validate_report`,
   and the profile runner binds the retained build pair, source manifest, gate
   receipts, frozen protocol, and managed scratch cleanup. The retained
   provider processes exited successfully and report identity matched. The
   original profile summary remains an incomplete historical parse: its old
   `strace -c` parser accepted only the one row carrying an error field and
   left no-error rows unparsed (parsed totals were 1 versus reported totals
   13,443, 127,427, and 2,332,877), while its old `perf -x,` parser treated
   field 3 as `running_percent` even though GNU perf places the percentage in
   field 4. The existing parser test used an explicit `errors=0` field and did
   not cover this blank-column format.

   The immutable r5 observer correction supersedes those parsed values without
   rerunning a workload. It reparses the retained raw files with exact
   row/call/error conservation: owned has 30 syscall rows and 13,443 calls,
   file-warm has 34 rows and 127,427 calls, and short-read has 30 rows and
   2,332,877 calls; each has one error and no unexplained line. It records the
   corrected perf field contract and 14 events per provider. The r5 receipt,
   gate, source snapshot, helper custody, and final7 helper inventory are all
   bound and pass the acceptance verifier. Zero-valued L1 event counters remain
   raw observations, not evidence that no loads or misses occurred. The
   observed `write` counts are whole-child trace/report serialization and
   cannot be read as Rust sink writes.

   The recovered owned `perf.data` export exits successfully without launching
   a workload, so it is optional observer attribution and not a rerun result;
   its recovery source differs from the retained workload snapshot. The
   corrected stack receipt has 1,382 statistical samples, including a
   23-sample publish-marker subset. Bare `run_sample` symbols and truncated
   callchains leave phase ownership explicitly ambiguous; only marker-visible
   subsets support scoped attribution, and unmatched samples cannot be called
   setup work. The current `_strace_failure_status` distinguishes
   launch/ptrace setup failures (`unavailable`) from target failures (`failed`),
   and its focused fixture tests pass; the retained run itself exercised only
   `available`. No profile result supports an operation-local timing,
   cache-miss, sink-byte, or optimization claim.

2. The warm terminal validator now binds `tmpdir` to the expected
   `TEMP/managed/<attempt>/<label>/tmp` root and requires the matching cleanup
   root. This earlier hardening finding is closed; pilot2 terminals exercise
   the exact path and cleanup receipt.

3. Warm `_vectors` and cold `_cell_stats` intentionally summarize a selected
   allocator scope: calls, allocated/deallocated bytes, region peak, and warm
   peak increment. The methods file states that failed-allocation calls, live
   endpoints, and historical peaks remain in every raw row and are not
   promoted to per-operation RSS. This is a documented analysis boundary, not
   a synthesized zero. Pilot2 raw rows retain those fields and pass row
   validation.

4. The warm allocator validator no longer imposes the invalid
   `deallocation_calls >= reallocation_calls` inequality. It keeps the valid
   `allocation_calls >= reallocation_calls` relation and full live/peak
   conservation. The realloc-without-standalone-free regression is covered by
   the final helper gate, so this earlier finding is closed.

5. The direct cold row validator remains weaker than the warm validator: an
   in-memory tamper check accepts changes to `live_bytes_after`,
   `allocated_bytes`, and peak fields. The acceptance verifier now closes this
   gap at the custody boundary by re-reading all accepted raw reports and
   enforcing strict types, live-byte conservation, peak bounds, reallocation
   bounds, and exact 60-formal/3-pilot allocator inventory. Cold allocator rows
   may support the accepted evidence only through that strict acceptance path;
   callers of `cold_measure.validate_report` still need the direct-validator
   hardening tracked as defense in depth.

The cold report deliberately has no executable identity field: the cold sample
terminal binds the launched role binary, build receipt, source, command, and
report path before `validate_report` checks the child report. That outer
custody chain is acceptable, but it must remain part of every cold acceptance
check. The current cold validator also enforces aligned range bounds, exact
sample inventories, and bounded aggregate invocation counts; the earlier
review concerns on those points are closed.

The selected `ENV_KEYS` projection is an explicit custody scope, not a claim
that every inherited host variable is frozen. `PATH`, `HOME`, `RUSTUP_HOME`,
loader variables, and other inherited inputs remain outside that projection;
the README and receipts describe it as selected-environment custody. The
shared 0484 CPU lock is likewise intentional historical-lane serialization, as
documented by the bundle. Neither is an acceptance blocker under the stated
scope.

These findings are based on the current source, formal/pilot receipts, and
independent raw recomputations, not on a forged retained artifact. The protocol
hashes, terminal inventories, raw receipt hashes, formal analyses, corrected
observer receipt, and helper custody were independently rechecked. The
accepted-evidence manifest is present; at the review checkpoint, only the
cleanup proof and subsequent seal remained. Profile attribution remains
optional and scoped to the corrected retained-data receipt.

## Warm validator remediation

The canonical warm validator in `measure.py` now closes the three warm
evidence gaps described above. The range adapter is checked against the
retained logical and physical traces call by call: logical calls, requested
bytes, returned bytes, short reads, offsets, and the configured maximum range
must conserve. For delayed and range-control arms, the validator derives the
expected delayed-call count, paced-call count, and transfer-delay nanoseconds
from the arm configuration and each physical returned range. Unpaced arms
must report zero service counters. This keeps the transport model separate
from the underlying source trace while making their relationship auditable.

Row validation now uses strict JSON types for operation and oracle fields,
requires exactly one commit operation, and binds patch evidence to the exact
untimed preflight scope `untimed_preflight_commit_patch_oracles`. Allocator
rows must conserve live bytes as
`live_after = live_before + allocated - deallocated`, account for reallocation
calls, and keep historical, region, and post-operation peaks ordered. The
existing exact-one materialization and positive bounded sink checks remain in
place.

Analysis no longer fabricates zeroes for cache fields absent from the Rust
report. It retains the emitted successful-load and expected-load vectors plus
the typed materialization oracle, and marks unavailable vectors as unavailable.
It also emits physical requested and returned media-overlap vectors alongside
the logical vectors; request-size histograms and media totals are derived only
from retained ranges.

The allocator analysis scope is now explicit: row custody validates all raw
allocator fields, while summaries emit selected call/byte totals, region peak,
and warm derived peak increment. A formal results record must preserve that
distinction for both warm and cold lanes.

The final7 helper custody gate reports 72 passing tests, including tampered
adapter totals and pacing, typed commit operations, allocator live-byte
conservation, realloc-without-free behavior, patch scope, unavailable cache
fields, physical media vectors, observer correction, and verifier custody.
The current helper inventory and gate source match the verifier. Warm and cold
pilot2 verification receipts pass over complete pilot inventories; these checks
do not authorize a performance claim.

## Acceptance sequence

The pilot and formal captures, analyses, verification receipts, terminal
inventories, corrected observer receipt, helper custody, and independent raw
recomputations are complete. The accepted-evidence manifest is assembled and
the strict acceptance-layer cold allocator audit is in place. The cleanup
proof has now been produced and validated; the final verifier and seal can
proceed. The optional profile's `strace`/`perf` parser receipt is complete,
but its whole-child scope and ambiguous stack qualification must stay explicit.
All conclusions remain descriptive same-binary provider and same-cold-cell
evidence, with no warm-versus-cold or before/after optimization claim.
