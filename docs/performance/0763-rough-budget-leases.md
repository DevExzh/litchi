# 0763 — rough budget leases take the streaming DOCX and XLSX writers from about a million budget atomics per large document to about 300, with every sole-holder refusal unchanged

Status: retained, implemented in `litchi-core`, `litchi-docx` and
`litchi-xlsx`. `performance_claim: none` — the paired timings, instruction
counts and budget-update counts below are evidence, not a registered claim.

OLE2 and OOXML remain the active priority; ODF stays deferred and iWork is
excluded. Base `6ec785c265` (the tip of record
[0762](0762-streaming-batch-compression.md) on the same branch,
`perf/0762-streaming-compression-and-leases`); commit `1ab98ca669`. The
coordinator's task: apply owner decision 6 of
[0758](0758-owner-decisions-2026-09-24.md) ("accept a rough budget lease rather
than a precise one") to the per-paragraph and per-row budget charges of the
streaming DOCX and XLSX writers, with fewer than 1,000 budget atomics per large
DOCX iteration, no limit ever exceeded, and every counter settling exactly.

**Result.** With both legs built by the identical command,
`docx_streaming_create` falls from 36.501 to 34.217 ms on the large corpus
(−6.03%, 95% CI [−7.15%, −5.79%]) and from 2.345 to 2.174 ms on the medium
one (−7.51%); `xlsx_streaming_create` from 90.976 to 89.928 ms (−1.24%) and
from 5.704 to 5.620 ms (−1.30%). The PPTX streaming writer, which makes no
`litchi-core` budget charge, and the two preservation controls do not move
beyond noise. A large DOCX iteration now makes 299 budget atomic updates
instead of 1,048,621, a large XLSX one 307 instead of 786,568. Every output
byte is unchanged, and a base-versus-candidate probe over 28,860 limit and
cancellation scenarios prints byte-identical transcripts: the same refusing
call, the same typed refusal and the same settled counters for a sole holder.
The gain is smaller than the flat profiles suggested (15–25% of the large
DOCX process's samples on `consume`): each removed atomic saved about 8
cycles. A locked update also absorbs stalls on the stores before it, which
the writer still pays without it, as
[0752](0752-streaming-writer-small-write-batching.md) observed; that is the
likely reason, not measured here.

## What was changed

* `crates/litchi-core/src/budget.rs`
  * `Budget::lease(resource, chunk) -> Lease`. Opening claims nothing.
  * `Lease::consume(amount)` hands out units the lease already holds without
    touching shared state; when it holds too few it claims
    `max(chunk, amount − held)` through the new private `claim_chain`, level by
    level, innermost first. Each level is one atomic check-and-add of the
    current grant shrunk to the room that level has left; a level with less
    room than the charge needs refuses, and the levels already charged are
    released before the refusal returns (as `charge_chain` does). When a level
    shrinks the grant, the levels already charged give the difference back, so
    every level ends up holding the same amount and no level is ever above its
    limit, even transiently. A refusal reports the refusing level, its limit
    and `used + need` — for the lease's own holder, exactly the value exact
    accounting of the whole charge reports.
  * `Lease::refund(amount)` takes back units handed out whose work did not
    happen (bounded by what the lease handed out) and returns whatever the
    lease would then hold beyond one chunk to the budget at once, so a lease
    never holds more than one chunk (added after review, below);
    `Lease::release` returns what the lease holds to the budget; `Drop`
    releases, on every path. `held`, `consumed`, `chunk` and `resource`
    report the lease's state.
* `crates/litchi-core/src/execution.rs`: `ExecutionContext::lease` and
  `ExecutionLease`, whose `consume` checks the context's cancellation token
  first, exactly as `ExecutionContext::consume` does, so replacing `consume`
  calls with lease charges keeps every cancellation check in place.
  `crates/litchi-core/src/lib.rs` re-exports `Lease` and `ExecutionLease`.
* `crates/litchi-docx/src/streaming.rs`: the writer holds `DocumentLeases` —
  Objects in 4 Ki chunks, Work in 64 Ki and input bytes in 64 KiB — opened
  after the fixed construction charges. The per-paragraph and per-run Objects
  and Work charges, the text scan's Work checkpoints and each text's input
  bytes go through them. The input bytes are still charged before any of the
  text's XML is emitted, and refunded where the old reservation was dropped
  (an emission failure). Every public call runs through a `settle` step that
  releases the leases once the writer is poisoned; `finish` releases them
  whatever its outcome; dropping the writer releases them.
* `crates/litchi-xlsx/src/streaming.rs`: the writer holds `WorkbookLeases` —
  Work in 64 Ki chunks and Objects in 4 Ki — opened after the construction
  charges. Row and cell Work, row Objects and the final part's Object go
  through them (the final Object must, or the writer's own unspent claim would
  count against it). A row's Objects are refunded on each path where the old
  reservation was dropped (sheet-XML limit, missing part, write failure).
  `write_row` releases the leases once the writer is poisoned; `finish`
  releases them whatever its outcome.
* `docs/adr/0005-io-memory-and-performance.md`: a dated clarification of "Every
  operation charges a hierarchical resource budget" under decision 6 (below).

Unchanged: the construction charges of both writers, the output-byte
reservations of their sinks, the scratch Memory reservations, `consume`,
`reserve`, `reserve_scoped` and every other budget user in the workspace. The
PPTX streaming writer makes no `litchi-core` budget charge (its "budget" is a
local output counter), so no lease applies there.

**Breaking changes: none.** The additions are additive. What other holders of
a shared budget observe does change, as decision 6 allows (below).

## Authority

Owner decision 6 of [0758](0758-owner-decisions-2026-09-24.md): budget
accounting may be chunked; an operation acquires a lease of up to a chunk and
consumes it locally; siblings observe the pre-claimed amount; a refusal may
come earlier than exact accounting at lease granularity — provided a limit is
never exceeded, consumption never exceeds the lease, unused lease returns on
drop including on error and cancellation, `ResourceLimit` errors stay typed,
the lease size is stated with its reason, and concurrency tests show every
counter settling exactly. Standing trade-offs of
[0652](0652-owner-decisions-for-the-third-wave.md): 2 (correctness first: no
limit is weakened; a sole holder's refusals do not move), 3 (the benign common
case — one writer per budget — is the one made cheap). ADR 0005 (hierarchical
budgets; the clarification below), ADR 0031 (its reservation dimensions are
admission-time reservations and are untouched), ADR 0002 (no new dependency;
`litchi-core` stays at the bottom).

**ADR wording.** ADR 0005 says "Every operation charges a hierarchical
resource budget supplied by an execution context". Read strictly, that could
require each operation to charge exactly what it uses when it uses it, so a
dated clarification under decision 6 is added: charging through a lease is
charging, and a lease never holds more than one chunk; what stays exact
(limits never exceeded, even transiently; the refusals of a holder whose
charges of a resource all go through one lease, and their values; refused
claims change nothing; unspent units return on release, drop, error,
cancellation and poison, so counters settle exactly) and what becomes rough
(other holders see pre-claimed units and can be refused early by up to one
chunk per other open lease). ADR 0031 has no sentence requiring
exact per-operation charging; its `Workers` and `IoConcurrency` permits are
reservations, not charges, and are not leased. Its paragraph is left as is.

## The evidence that motivated it

After [0752](0752-streaming-writer-small-write-batching.md) removed the
reference counting around each charge, the seven `consume` charges per DOCX
paragraph (plus one input-byte reservation) were "the largest item of the DOCX
profile (26% of samples)", each one atomic update that "exact, shared-budget
semantics require". The coordinator's profile attributed 13.7% of the XLSX
large iteration to budget accounting
(`results/change-0756/profile-r2/REPORT.md`). After 0762, a flat profile of
the large DOCX case puts `ExecutionContext::consume` first, at 15.3% of the
whole process's samples, and of the large XLSX case at 9.1% (plus
`reserve_scoped` 1.6%) (`results/change-0762/profiles/*-after.top.txt`).

## Lease sizes and why

| writer | resource | chunk | per large iteration: charges → claims |
| --- | --- | ---: | --- |
| DOCX | Work | 64 Ki units | 6,422,528 units in 655,360 charges → 98 claims |
| DOCX | input bytes | 64 KiB | 6,029,312 bytes in 131,072 charges → 92 claims |
| DOCX | Objects | 4 Ki | 262,144 in 262,144 charges → 64 claims |
| XLSX | Work | 64 Ki units | 655,360 in 655,360 charges → 10 claims |
| XLSX | Objects | 4 Ki | 655,360 in 131,072 charges → 160 claims |

A chunk bounds two things: the number of shared updates (the charge volume
divided by the chunk) and the most another holder of the same budget can see
pre-claimed per writer and resource (one chunk). 64 Ki Work units and 64 KiB
of input are under 0.01% of every finite profile's Work and input ceilings
(`Profile::Server`: 10⁹ Work units, 2 GiB input); 4 Ki Objects are 0.04% of the
server profile's 10⁷ objects. Near a limit a claim shrinks to the room left,
so a sole holder never needs more than its charge.

## What stays exact and what becomes rough

For the lease's own holder, nothing moves, provided every charge it makes of
the resource goes through that one lease; a charge it made outside the lease
would see the lease's unspent units as used. Both writers meet this: their
construction charges are made before their leases open, and every later
charge of a leased resource goes through its lease. A charge the lease covers
touches no shared state. A charge it cannot cover needs `need = amount −
held` more; every level already holds the lease's `held` units, so a level
with room `limit − used ≥ need` is exactly a level where exact accounting has
room for `amount`, the refusing level is the innermost one exact accounting
would name, and the refusal's `used + need` equals exact accounting's
`used_exact + amount`. When the writer is poisoned or finishes, it releases
what it holds, so every counter then equals exactly what was handed out.

For other holders of a shared budget (sibling budgets under a common parent,
or several operations on one budget), the units a lease holds and has not
handed out count as used: a sibling can be refused earlier than exact
accounting would refuse it, by up to one chunk per open lease of another
holder, and its refusal reports usage that includes the pre-claims. A lease
never holds more than one chunk: a claim leaves it less than one chunk, and a
refund returns whatever would exceed one chunk. No level is ever over its
limit: every claim is a check-and-add, and a shrinking claim only gives back.

## Measured

Host AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`, shared with other agents;
every measured process pinned to CPU 12 with `taskset`; no build of this
record ran during a measurement. Harness `tools/perf-baseline`, unchanged.
Both legs are built by the identical command from their own trees: before from
the detached worktree `0762-before-src` at `6ec785c265`, after from the branch
at `1ab98ca669` (`binaries.txt`), and copied to paths of equal length.

### Timing

Four rounds; in each, every case ran before, after, after, before, so eight
processes per leg, paired within the round (slot 1 with 2, slot 4 with 3). The
table shows the medians of the process p50s and p95s, the median paired p50
change and a percentile bootstrap 95% interval over the eight paired changes
(20,000 resamples, seed 762). Raw reports: `timing/abba-raw.tar.gz`; summary:
`timing/summary.json`.

| case, corpus | processes × samples | before p50 ms | after p50 ms | paired p50 change | 95% CI | p95 ms, before → after |
| --- | --- | ---: | ---: | ---: | --- | --- |
| `docx_streaming_create` large | 8+8 × 20 | 36.501 | 34.217 | **−6.03%** | [−7.15%, −5.79%] | 36.983 → 34.366 |
| `docx_streaming_create` medium | 8+8 × 40 | 2.345 | 2.174 | **−7.51%** | [−8.12%, −6.82%] | 2.442 → 2.213 |
| `xlsx_streaming_create` large | 8+8 × 15 | 90.976 | 89.928 | **−1.24%** | [−1.63%, −1.04%] | 91.421 → 90.498 |
| `xlsx_streaming_create` medium | 8+8 × 30 | 5.704 | 5.620 | **−1.30%** | [−2.51%, −1.08%] | 5.831 → 5.638 |
| `pptx_streaming_create` large | 8+8 × 15 | 168.558 | 167.917 | −1.05% | [−2.40%, +0.56%] | 172.302 → 170.672 |
| `pptx_streaming_create` medium | 8+8 × 30 | 5.916 | 5.874 | −0.58% | [−1.49%, +0.69%] | 6.052 → 5.997 |
| control: `docx_semantic_one_edit_save` large | 8+8 × 40 | 5.667 | 5.533 | −0.47% | [−2.93%, +0.41%] | 5.875 → 5.644 |
| control: `xlsx_ordinary_save_lifecycle` | 8+8 × 20 | 14.251 | 14.392 | +0.05% | [−2.80%, +2.38%] | 14.500 → 15.421 |

Every process of both legs publishes the same package per case: `e263d219…`
(DOCX large), `04e3ea69…`, `edbf573c…`, `7a708509…`, `8658e17d…`, `6cfa2234…`
and `20335c44…` (the ordinary-save control); the semantic-edit control reports
no digest. PPTX is not touched by this record; its −1.05% (CI spanning zero)
is placement and noise.

### Regressions and every over-5% flag

No case regresses beyond noise. Of the 56 paired comparisons over 5% at p50,
p95 or mean (`timing/flags.json`), 51 are favourable (45 on the two DOCX
cases). The 5 adverse ones are all on the ordinary-save control, whose code
this record does not touch: rounds 2 and 3, p95 +5.9% to +19.9% and mean
+5.9% and +7.3%, one tail process each; the control's paired p50 median is
+0.05% and its instructions are unchanged (below).

### Instructions and cycles

User-space counters per timed iteration, by differencing a low- and a
high-sample run of the same binary (`scripts/perf_delta.py`); two rounds of
before, after, after, before; medians of four measurements per leg and the
median paired change (`counters/summary.json`):

| case | instructions before → after | change | cycles before → after | change |
| --- | --- | ---: | --- | ---: |
| DOCX large | 882.47M → 878.47M | −0.30% | 162.60M → 154.44M | −5.04% |
| DOCX medium | 56.01M → 55.70M | −0.60% | 11.96M → 9.94M | −18.84% |
| XLSX large | 1,966.87M → 1,958.23M | −0.44% | 401.18M → 403.03M | +1.32% |
| XLSX medium | 122.37M → 121.83M | −0.44% | 25.85M → 25.50M | +1.52% |
| PPTX large | 2,695.21M → 2,703.22M | +0.30% | not interpretable | — |
| PPTX medium | 93.29M → 93.29M | −0.00% | 27.52M → 26.65M | −3.06% |
| control: semantic edit | 349.02M → 348.95M | −0.03% | 72.49M → 70.03M | −2.97% |
| control: ordinary save | 74.13M → 74.14M | +0.01% | 26.97M → 26.63M | −2.23% |

The leases remove few instructions — a covered charge is a comparison and a
subtraction instead of a cancellation check and a compare-exchange loop — and
the atomic updates' cost was in cycles. On the large DOCX iteration the 1.05
million removed updates are worth 8.2 M cycles (about 8 cycles each), close to
the 2.3 ms wall-clock gain. Two-run cycle differences on this shared host are
noisy (the medium DOCX pairs range from −8% to −36%, the XLSX ones from −10%
to +10%); PPTX large's are lost in its 120 G-cycle corpus reopen, and its
instruction measurements agree within about 1% per leg, as in
[0762](0762-streaming-batch-compression.md).

### Allocations

Allocator counters per timed iteration from the allocation build
(`alloc/summary.json`): every streaming case and the ordinary-save control
make the same number of allocations of the same total size on both legs
(131,114 for DOCX large, 131,145 for XLSX large, 278,156 for PPTX large;
region peaks within 2 bytes). A lease holds a shared handle to its budget node
and allocates nothing. The first allocation run stopped after its first report
without an error message; the whole allocation run was repeated once and
completed.

### Budget atomics per large iteration

`atomic_count` (`probe/atomic_count/`, built against the candidate) runs the
harness's large DOCX and XLSX scripts against a single root budget and reads
every counter after every call. With one holder and one level, every claim or
release is one atomic update that changes its counter, and a covered charge
changes nothing; output bytes are reserved once per sink write, counted with a
counting sink (`probe/atomic-count.txt`):

| writer, large | leased counters: claims and releases | output reservations | fixed (scratch, construction) | total | exact accounting |
| --- | ---: | ---: | ---: | ---: | ---: |
| DOCX (131,072 paragraphs) | 254 (64 Objects, 98 Work, 92 input) | 41 | 4 | **299** | 1,048,621 |
| XLSX (131,072 rows) | 172 (162 Objects, 10 Work) | 126 | 9 | **307** | 786,568 |

The DOCX totals are exact multiples of their chunks, so no final release is
needed there. The XLSX rows' Objects are too; `finish` then claims the room
left for its one Object and releases the rest in the same call, which the
probe's between-call reading sees as one change (161 printed), so the table
counts both. "Fixed" is the scratch reservation and its release and the
writers' construction charges, which stay exact. Exact accounting is 8 charges
per DOCX paragraph (2 Objects, 5 Work, 1 input reservation) and 6 per XLSX row
(5 Work, 1 Objects reservation), plus the same output, finish and fixed
updates. Both lease totals are under the task's 1,000.

## Correctness evidence beyond the tests

* **Base against candidate, a sole holder.** `lease_diff`
  (`probe/lease_diff/`), built from one source against each tree, drives the
  DOCX writer (24 paragraphs, some with two runs or two text calls, 140 calls)
  and the XLSX writer (40 rows of 0–4 cells) through 28,860 scenarios: every
  value of every budget dimension (Work, Objects, input and output bytes; the
  Memory reservation at 0, 1 and one either side of its size) from zero to one
  past the script's total, with the limit on the writer's own root and on the
  parent of its budget, and a cancellation before each call. For each scenario
  it prints the refusing call and the complete error (resource, observed,
  limit, output bytes written) or the published package's SHA-256, and every
  level's usage of every resource once the writer has finished or been
  dropped. The two transcripts are byte-identical (`90f8efd8…`,
  `probe/lease-diff-summary.txt`: 7,300 Work, 13,212 input-byte and 104 Objects
  refusals at DOCX calls, 240 Work and 224 Objects refusals at XLSX rows, the
  refusals at construction and at finish, 181 cancellations, 41 published
  packages).
* **Output bytes do not move.** Every harness process of both legs publishes
  the same package per case (below).
* **Mutation checks.** Removing the rollback of a refused claim fails four
`litchi-core` lease tests, including the sole-holder sweep and the concurrency
test. Replacing the check-and-add with an add followed by a check fails the
concurrency test in three runs of three (its monitor or a worker's refusal
assertion). Removing the release from the DOCX writer's poison step fails the
DOCX poison test and the limit sweep. Each mutation was reverted and the file
compared equal.

## What is not claimed

No claim is registered. The numbers are scoped to the named synthetic corpora,
this host and CPU 12. They do not establish:

* the behaviour of several writers sharing one budget under load beyond what
  the unit tests show: the harness contexts are single-level roots with one
  writer;
* anything about budget users other than the two streaming writers, which
  still charge exactly;
* the effect of the pre-claim on a caller that shares a small budget between
  writers; such a caller can be refused early by up to one chunk per other
  open lease, as decision 6 accepts.

## Verification

All gates pass at `1ab98ca669` except the known base failure listed below
(`results/change-0763/gates.txt`): `cargo fmt --all --check`;
`cargo check --all-targets` of `litchi-core` and every workspace package that
depends on it (40 with the facade and `litchi-py`, iWork excluded; no
warnings); warning-denied Clippy on the library and all targets of
`litchi-core`, `litchi-docx` and `litchi-xlsx`; warning-denied rustdoc of the
three; `cargo test` of `litchi-core` (238), `litchi-docx` (1,881),
`litchi-xlsx` (2,039), the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`
(382), 24 other OOXML and OLE2 dependents of `litchi-core` and, as an extra,
its 11 ODF dependents; crate boundaries; the structural claims check (10
claims). `tools/non_iwork_gate.py verify` fails with "unexpected:
litchi-xldm", as at the base (a known failure being fixed separately). Test
builds used `CARGO_PROFILE_DEV_DEBUG=0` and `CARGO_INCREMENTAL=0` to keep the
target directory small on the shared disk. The harness is unchanged; its tests
and the coverage validator were not run.

New tests. `litchi-core` `budget.rs`:
`a_lease_claims_chunks_and_hands_units_out_locally`,
`a_sole_lease_holder_is_refused_exactly_as_exact_accounting_refuses` (40
charges of 1–9 units under every limit from 0 to one past their total, with
the tight limit on the leaf, the middle level, the root, or leaf and middle,
for chunks 1, 2, 5, 16, 64, 1,000 and `u64::MAX`: the same refusing charge,
the same `ResourceLimit` and, after release, the same counters as `consume`;
every intermediate state has equal counters at every level and none over its
limit), `siblings_see_pre_claimed_units_and_no_level_passes_its_limit`,
`a_claim_near_a_limit_shrinks_to_the_room_at_every_level`,
`refunds_return_units_to_the_lease_and_release_settles_every_level`,
`a_lease_releases_when_dropped_even_while_unwinding` (a panic and an error
path), and `concurrent_leases_and_reservations_keep_a_shared_hierarchy_exact`
(eight threads on the six-node tree of 0752's reservation test mix lease
charges, refunds, releases and re-opened leases of random chunk sizes with
exact `consume` and Memory reservations; a monitor thread reads every level
throughout and never sees one above its limit; afterwards every level's Work
equals exactly what its subtree was granted and not refunded, and Memory is
zero). `execution.rs`:
`lease_charges_check_cancellation_first_and_charge_the_shared_hierarchy`,
`lease_refusals_are_the_exact_typed_resource_limit`, and the Send/Sync check
extended to `Lease` and `ExecutionLease`. `litchi-docx` `tests/streaming.rs`:
`a_sole_writer_is_refused_exactly_at_every_objects_work_and_input_limit`
(every limit from the construction charge to one past the total, on each of
Objects, Work and input bytes: the refusing call, resource, observed and limit
of a model of the documented accounting, and the settled counter after the
poisoned writer returned its leases),
`a_sibling_sees_the_writers_pre_claim_until_the_writer_finishes`, and
`leases_are_returned_when_the_writer_is_poisoned_or_dropped`. `litchi-xlsx`
`streaming.rs`: `a_sole_writer_is_refused_exactly_at_every_objects_and_work_limit`,
`a_sibling_sees_the_writers_pre_claimed_objects_until_it_finishes`, and
`leases_are_returned_when_the_writer_is_poisoned`.

## Review corrections

An independent review (its probes are not part of this packet) confirmed the
differential and the limits: 16 threads on a seven-node tree with 342,812
refusals never saw a level over its limit, 3,000 random sole-holder cases
matched exact accounting, and base-against-branch refusal transcripts of both
writers (11,386 DOCX and 1,247 XLSX scenarios) were identical. It found one defect, fixed in `3f4e7f5d71`:

* **A refund was not capped.** `Lease::refund` returned the units to the
  lease whatever their amount. The XLSX writer refunds a row's objects when
  the worksheet-XML limit refuses the row, which does not poison the writer,
  so after a refused 12,000-cell row the lease held 12,001 objects — about
  three chunks — until `finish` or drop (the base held none). No limit was
  exceeded, but decision 6 grants a lease of up to one chunk. A refund now
  returns whatever the lease would hold beyond its chunk to the budget and
  every ancestor at once; `consume` already kept the lease within one chunk.
  New tests: `a_refund_wider_than_a_chunk_returns_the_excess_at_once`;
  `a_lease_never_holds_more_than_one_chunk` (20,000 random calls per chunk
  size of 1, 3, 64, 4,096 and `u64::MAX` — charges, charges wider than a
  chunk, accepted and refused refunds, releases — with refusals under a
  three-level limit: after every call the lease holds at most one chunk and
  every level shows exactly what was handed out plus what the lease holds); the concurrency test now
  checks the cap after every step; and the XLSX
  `a_refused_row_leaves_at_most_one_chunk_pre_claimed` (rows of 150 to 12,000
  cells refused by a 4 KiB worksheet-XML limit, each leaving the writer
  usable, at most one chunk held and the budget exact, then settling exactly
  at `finish`). Without the cap three `litchi-core` tests and the XLSX test
  fail; the XLSX test then shows 12,001 objects held.
* **Wording.** The `Lease` rustdoc, the ADR 0005 clarification and this
  record now state that a holder's refusals are exact when all its charges of
  the resource go through one lease, and that another holder can be refused
  early by up to one chunk per other open lease.

The fix touches only refund paths, which the measured runs never take; the
timings, counts and transcripts above stand and were not re-measured. After
the fix (and 0762's review corrections in `7b0bd40c02`), with a fresh target
directory: `cargo fmt --all --check`; warning-denied Clippy on the library
and all targets and warning-denied rustdoc of `litchi-core`, `soapberry-zip`,
`litchi-docx` and `litchi-xlsx`; `cargo test` of `litchi-core` (240),
`soapberry-zip` (654), `litchi-docx` (1,881), `litchi-xlsx` (2,040) and, as an
extra, `litchi-opc` (918) and `litchi-pptx` (1,167) — all pass
(`results/change-0763/gates.txt`).

## Cleanup

Binary digests are in `binaries.txt`. After the evidence was copied, the
target directories of both records (`targets/0762`, `targets/0763`,
`targets/0763-before`, the two `lease_diff` probe targets), the detached before
worktree `0762-before-src` (`git worktree remove --force`, then `git worktree
prune`) and the scratch directory `scratch/0762` were removed
(`cleanup.json`). The branch worktree is kept.
