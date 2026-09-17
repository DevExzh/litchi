# 0662: a publication's changed members deflate in parallel under the caller's execution budget, 1.21-1.52× on real multi-member saves and bit-identical at every width

Status: retained, implemented for the managed preservation-write boundary.
`performance_claim: none` — the paired medians and counts below are reported as
evidence, not registered as claims. Accepted ADR 0031's ZIP-read, CFB-read and
source-backed-read composition rows remain outside this record and are named in
the limitations below; this record does not present the write bridge as full
ADR 0031 completion.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `ab07e2a47` (change 0654); branch `perf/0662-parallel-changed-member-deflate`.

## What was changed

The preservation writer's compression of a publication's *changed* members
— `PreservationIndex::prepare` calling `generated_entry` once per
`PreservationAction::Regenerate` — can now run as a bounded wave of concurrent
tasks instead of a serial loop. Three crates move:

* **`litchi-core`** gains the two execution budget dimensions this wave is
  charged against (`Resource::Workers`, `Resource::CpuTasks`), the per-task
  floor of [ADR 0031](../adr/0031-execution-context-budgets.md) §7
  (`ExecutionLimits::min_task_bytes`), and the caller-provided executor of
  ADR 0031 §4 (`ScopedWorkers`, attached to an `ExecutionContext`). It creates
  no thread and depends on no scheduler.
* **`soapberry-zip`** gains `office::ParallelWriteSession`, the write-side
  counterpart of `ParallelReadSession`: a session that owns at most one local
  pool, builds it **lazily** on the first wave that qualifies, and takes a
  caller's facility instead when one is attached. `PreservationPlan::deflate_wave`
  computes the balance rule from the plan alone, and two new entry points,
  `PreservationIndex::write_to_with_session` and
  `write_to_with_accounting_and_session`, publish under it. The crate gains no
  dependency: it defines its own `ScopedWorkers`, exactly as it already defines
  its own `CancellationProbe`.
* **`litchi-opc`** bridges the two on the *managed* source-backed publication
  route, the only write path in the workspace that carries an
  `ExecutionContext`. A managed package builds one inert session; each
  publication asks it for the wave its changed set earns, charges that wave
  against the caller's budget, and publishes under it.

`write_to` and `write_to_with_accounting` keep their signatures and their
behaviour; an unmanaged package, and a managed one whose wave is one task wide,
take exactly the loop they took before, build no pool and start no thread.

### The balance rule, as implemented

Computed from the plan before any member is compressed:

1. at least two regenerated members are Deflate compression tasks — a
   `Precompressed` member is framed from bytes that are already compressed and
   a `Store` member is copied, so neither is a task (change 0593's pristine
   members were never candidates at all: they are `Copy` actions);
2. the **remainder** — total task payload minus the largest task — is at least
   the policy's threshold. The remainder, not the total, is the predictor,
   because the largest member is the pole at every width;
3. the granted width is `min(workers, in-flight tasks, task count)`, narrowed
   further by `1 + remainder / min_task_bytes` when the caller sets a per-task
   floor;
4. the threshold in rule 2 depends on whether the session already holds its
   workers: `DEFAULT_MIN_PARALLELIZABLE_BYTES` (8 KiB) when a caller facility
   or an already-built pool is in hand, `DEFAULT_MIN_POOL_PARALLELIZABLE_BYTES`
   (64 KiB) when the session must build a pool first, because that wave has to
   repay the pool as well as itself.

Both thresholds were measured on this base (§"Measured", the balance-rule
sweep), not inherited from change 0624's isolated-deflate figures.

## Authority

Decision 6 of [0652](0652-owner-decisions-for-the-third-wave.md), in the owner's
words: **"Parallel deflate: accept ADR 0031."** 0652 records what that
authorizes — "ADR 0031 is Accepted as of this record; parallel deflate of
changed members under the execution context's budgets, with the measured
balance rule of 0624" — and what this record must still prove: "no hidden
global pool (ADR 0005); the balance rule measured, not assumed; determinism of
the published bytes (0625, 0631) unchanged; the documented ordinary save's
publication-bound nature (0638) stated so the saving is claimed only where it
exists." Each is answered below.

[0624](0624-parallel-changed-member-deflate-design.md) froze the design this
implements, at the boundary it named (§4.1), with the streaming writer
explicitly out of scope (§4.2), the execution-context shape of §4.3, and the
threshold rule of §4.4 re-measured here as §4.4 required. Trade-off 1 of 0652
("breaking changes are totally acceptable") authorizes the public API additions
listed below.

## Breaking changes

`litchi-core`:

* `Resource` gains `Workers` and `CpuTasks`. The enum is `#[non_exhaustive]`,
  so no downstream `match` breaks; in-workspace matches already carry a
  wildcard. The private `RESOURCE_COUNT` moves 6 → 8, so every `Budget` node
  grows by two `AtomicU64`.
* `Limits::new` keeps its six arguments and leaves the two new dimensions
  `u64::MAX`, so all of its call sites keep exactly their current meaning. New:
  `Limits::with_execution(workers, cpu_tasks)`.
* `Limits::for_profile` now names finite values in the two new dimensions, as
  ADR 0005 requires production profiles to be finite: `Workers` 64 / 256 / 1024
  and `CpuTasks` 1e6 / 5e6 / 5e7 for `Server` / `Desktop` / `TrustedBatch`.
  Workers follows the same ×4 ladder the other dimensions use and is far above
  the 20 threads change 0615 measured at width 8; `CpuTasks` is `Work / 1000`,
  one task per declared kilobyte, against a corpus whose largest changed set is
  59 tasks.
* `ExecutionLimits` gains a private `min_task_bytes` field with
  `with_min_task_bytes()` and `min_task_bytes()`. `new` and `with_affinity`
  keep their signatures and default the floor to zero. The field participates
  in the derived `PartialEq`/`Eq`/`Hash`.
* `ExecutionError` gains `TaskFloorExceedsInFlightBytes { min_task_bytes,
  max_in_flight_bytes }`. The enum is `#[non_exhaustive]`.
* New trait `ScopedWorkers`, and `ExecutionContext::with_scoped_workers()` /
  `scoped_workers()`. `ExecutionContext` gains a private field; it is
  constructed only through `new`, which is unchanged.

`soapberry-zip`:

* New in `office`: trait `ScopedWorkers`, `SerialScopedWorkers`,
  `ParallelWriteLimits`, `ParallelWriteSession`, and the constants
  `DEFAULT_MIN_PARALLELIZABLE_BYTES`, `DEFAULT_MIN_POOL_PARALLELIZABLE_BYTES`,
  `DEFAULT_MIN_TASK_BYTES`, `WORKER_STATE_BYTES`.
* New at the crate root: `DeflateWave`, `PreservationPlan::deflate_wave`,
  `PreservationIndex::write_to_with_session` and
  `write_to_with_accounting_and_session`.
* `ErrorKind` gains `InvalidParallelWriteLimits`, `ParallelWriteWorkerPool` and
  `ParallelWriteWorkerPanic`. The enum is `#[non_exhaustive]`.

`litchi-opc`: no public item changed. The bridge, the session and the permits
are private to `SourceBackedPackage`'s cache.

No ordinary CRUD signature names any of these types: `ScopedWorkers`,
`ParallelWriteSession` and `DeflateWave` appear only on `ExecutionContext`, on
`soapberry-zip`'s own advanced API, and inside `litchi-opc`. The audit is in
§"Correctness evidence".

## Why it is sound

**The published bytes cannot depend on the order members are compressed in.**
`generated_entry` is a pure function of one `&RegeneratedEntry`: it borrows
nothing from the index, touches no shared state, and returns an owned
mini-archive. Results are placed in fixed plan slots, and emission order is
`self.local_order` (source physical order) and the entry index
(central-directory order) — the writer's own documented contract that "the
action list is not an ordering control". Scheduling therefore moves *when* a
member is compressed and nothing else. That is the argument; §"Measured" and
§"Correctness evidence" are the proof, over 305 real packages at four widths.

**Nothing is emitted until every member is compressed.** `write_prepared` calls
`prepare` first and writes the first local record only afterwards — the same
fence the serial writer has always had. A wave that is cancelled, that
exhausts a budget, or whose member is refused therefore leaves the sink
untouched, which is what keeps change 0497's atomic publication atomic.

**Error identity and error *order* are the serial ones.** The serial loop
returns the first error in plan-action order, not the first in time. The wave
collects every task's `Result` into its plan slot; the ordered walk then takes
each slot in plan order and returns the first failure it reaches, exactly where
the serial loop would have returned it. A plan whose second and fourth members
are both refused returns the second member's error at every width, byte for
byte of message text, with the sink untouched — a retained test.

**Cancellation is observed between members, never inside one.** A task that has
started always finishes, because `docs/GOAL.md` rule 3 forbids abandoning
started work for a partial result; the probe is read before the wave, before
each member, and after the wave, and the error a cancelled publication returns
is the post-wave probe's, so it is deterministic rather than a race between
tasks.

**No hidden pool, and no ambient thread.** The session owns at most one Rayon
pool, built lazily on the first qualifying wave and dropped with the package.
Rayon's global pool is never initialized or installed. When a caller attaches a
`ScopedWorkers` facility, the session builds no pool at all and every task runs
on the caller's executor — proven by a test that counts the session's pool
threads at zero and the facility's waves at one.

**The budget bounds the wave, and refuses before any work.** A wave charges one
`Resource::CpuTasks` unit per task at every width, and a wave wider than one
reserves `Resource::Workers` permits and the memory its workers retain, before
the first member is compressed. A budget that grants fewer permits than the
policy asked for **narrows** the wave and proceeds, as ADR 0031 §2 requires; a
budget that cannot afford the tasks refuses with the existing
`ResourceLimit { resource: CpuTasks, .. }` shape and an untouched sink. Permits
are held for the life of the pool they admitted and released with the package.

**0618's allocator finding is untouched.** Change 0618 rejected a compressor
carried across members because the plan retains one buffer per regenerated
member, so the compressor could not be recycled and cost 271 minor faults per
publish. This change keeps 0618's decision exactly: one `ReusableDeflateState`
per member, constructed and freed inside it. There is no per-worker compressor
reuse to re-measure, and the measured minor faults per publication *fall* (§
"Measured").

**Contracts that did not move.** Every limit, refusal, audit and fence is where
change 0654 left it. The compressed-size budget stays per write call inside
`LimitedEntryWriter`; the mini-archive structural checks stay inside
`generated_entry`; `verify_authored`, the publication plan, the original-bytes
audit and the signature policy are untouched; the streaming writer
(`StreamingArchiveWriter`), which 0624 §4.2 placed out of scope, is not
modified, so the authored and eager `PackageWriter` routes are unchanged.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; `rustc 1.95.0`;
`cargo build --release --locked` (workspace `[profile.release]` has `lto = true`
and `panic = "abort"`; the scratch probe adds `debug = 1`). Twelve other agents
were building and measuring throughout; run-window load average 10-40.

The brief assigns this change CPU 17. One CPU cannot measure a width-8 wave at
all, so the scaling legs are pinned to CPU 17 plus the seven quietest other
CPUs, chosen from `/proc/stat` immediately before each run and recorded with
it; the A/A floor is measured in the same window on the same set. Every
sequential and width-1 leg is the same single-threaded work, so the wider mask
does not favour them.

### The A/A floor

Each table below carries its own floor, measured as the spread between the two
halves of the paired baseline leg in the same window. In the reported windows
they are 0.04%-1.07% on the ZIP sweep, 0.04%-0.83% on the real fixtures, and
0.12%-37.42% on the harness cases. Where a delta is inside its floor it is
reported as inside the floor and nothing is claimed from it.

### 1. The opt-in default costs nothing (`timing/harness-ab.tsv`)

`tools/perf-baseline` at `ab07e2a47` and at this branch, 30 samples per leg,
legs ordered before / after / after / before, `taskset -c 17`. Sixteen save
cases, including the 0638 ordinary-save family and the managed XLSX selectors
(which construct their context with one worker, `lib.rs:44497`):

| case | before p50 (ns) | after p50 (ns) | Δ p50 | A/A floor | output sha256 |
| --- | ---: | ---: | ---: | ---: | --- |
| `xlsx_eager_cell_values_one_edit_save` | 40,464,257 | 39,775,720 | −1.70% | 0.32% | identical |
| `xlsx_source_backed_cell_values_one_edit_save` | 28,981,174 | 29,200,176 | +0.76% | 1.82% | identical |
| `xlsx_source_backed_cell_values_one_percent_edit_save` | 32,798,703 | 33,056,721 | +0.79% | 1.09% | identical |
| `xlsx_source_backed_managed_cell_values_one_edit_save` | 29,120,635 | 29,298,787 | +0.61% | 1.11% | identical |
| `xlsx_source_backed_managed_cell_values_one_percent_edit_save` | 32,674,113 | 32,904,801 | +0.71% | 1.17% | identical |
| `docx_source_backed_one_edit_save` | 4,602,931 | 4,623,103 | +0.44% | 2.29% | identical |
| `pptx_source_backed_one_edit_save` | 7,120,782 | 7,206,346 | +1.20% | 37.42% | identical |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 15,299,410 | 12,439,422 | −18.69% | 24.44% | identical |
| `opc_source_overlay_one_part_save` | 62,401,644 | 62,483,292 | +0.13% | 0.26% | identical |
| `opc_source_overlay_multi_part_changed` | 6,161,802 | 6,076,227 | −1.39% | 3.17% | identical |
| `docx_ordinary_save_lifecycle` | 907,185 | 909,255 | +0.23% | 1.26% | identical |
| `docx_ordinary_save_atomic_publish` | 382,752 | 378,262 | −1.17% | 0.12% | identical |
| `xlsx_ordinary_save_lifecycle` | 8,601,544 | 8,580,332 | −0.25% | 4.55% | identical |
| `xlsx_ordinary_save_atomic_publish` | 2,218,302 | 2,230,841 | +0.57% | 2.75% | identical |
| `pptx_ordinary_save_lifecycle` | 2,663,444 | 2,647,734 | −0.59% | 1.52% | identical |
| `pptx_ordinary_save_atomic_publish` | 111,640 | 112,050 | +0.37% | 1.42% | identical |

Every case's `output_sha256` is the same on both legs and in both repeats.
Every p50 delta is inside its own A/A floor except `docx_ordinary_save_atomic_publish`
(−1.17% against a 0.12% floor, a p50 improvement this change has no mechanism to
produce and does not claim) and the two source-backed XLSX cases at +0.76% and
+0.79% against 1.82% and 1.09% floors, both inside. The −18.69% on the
media-rich cross copy sits against a 24.44% floor and is not a signal.

**Where the saving is and is not.** 0638 found the documented ordinary save
publication-bound, and the four `*_ordinary_save_*` rows above are the direct
consequence: that route runs `PackageWriter` without an `ExecutionContext`,
carries no session, and cannot take a wave at any setting. **No saving is
claimed for it.** The same holds for every unmanaged source-backed editor, for
the eager route, and for the decoded splice/replay route, whose plans are
`copy_all` and contain no regenerated member at all.

### 2. The balance rule, measured at the ZIP publication boundary (`scaling/zip-sweep.tsv`, `scaling/zip-sweep-tiny.tsv`)

Both thresholds set to zero so the crossover can be found rather than assumed;
one session reused across samples, i.e. the workers are already in hand; 60
samples per leg (80 for the tiny shapes), 10 warmups, paired sequential legs.

| changed set | remainder | sequential p50 | w1 | w2 | w4 | w8 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 2 × 64 B | 64 B | 10.17 µs | 1.008 | **0.648** | 0.532 | — |
| 2 × 512 B | 512 B | 14.19 µs | 0.983 | **0.628** | 0.584 | — |
| 4 × 256 B | 768 B | 25.64 µs | 1.016 | **0.753** | 0.707 | — |
| 2 × 871 B | 871 B | 15.98 µs | 1.003 | **1.476** | 1.484 | 1.493 |
| 4 × 871 B | 2.6 KB | 31.48 µs | 0.996 | 1.658 | 1.937 | 1.996 |
| 8 × 871 B | 6.1 KB | 63.10 µs | 1.002 | 1.768 | 2.321 | 2.017 |
| 16 × 871 B | 13.1 KB | 128.13 µs | 1.003 | 1.468 | 2.574 | 2.923 |
| 40 × 871 B | 34.0 KB | 341.47 µs | 0.999 | 1.910 | 3.108 | **4.307** |
| 128 × 871 B | 110.6 KB | 1,347.42 µs | 1.002 | 2.155 | 3.947 | **6.063** |
| 8 × 4,096 B | 28.7 KB | 143.45 µs | 1.000 | 1.879 | 2.894 | 3.103 |
| 40 × 4,096 B | 159.7 KB | 1,070.51 µs | 1.004 | 2.282 | 4.334 | 6.738 |
| 2 × 26,698 B | 26.7 KB | 212.10 µs | 1.004 | **2.005** | 1.994 | 2.003 |
| 16 × 26,698 B | 400.5 KB | 2,389.80 µs | 0.998 | 2.067 | 3.993 | **7.358** |
| 2 × 1,893,450 B | 1.89 MB | 22,653.00 µs | 1.005 | **2.006** | 2.010 | 2.011 |

Read three things from this table.

* **Width 1 is the sequential path**, everywhere: 0.996-1.008, inside every
  floor. No pool is built and no thread starts, which a test also asserts by
  counting the session's pool threads.
* **The crossover is a remainder of about 800 bytes**, not 256 KiB. A 768-byte
  remainder loses (0.75×); an 871-byte remainder wins (1.48×). The fixed cost
  of a wave on warm workers is about 9 µs.
* **The bound is the largest member.** The two-member 26.7 KB and 1.89 MB sets
  reach 2.00× at width 2 and go no further at 4 or 8, which is the arithmetic
  maximum for two equal poles; change 0624's model predicted 2.00 and measured
  1.988 for the same 1.89 MB pair.

`DEFAULT_MIN_PARALLELIZABLE_BYTES` is set to **8 KiB** — an order of magnitude
above the measured crossover. Every shape at or above it returned at least
1.77× at width 2 in this window.

### 3. The cold-pool threshold, measured on the OPC publication route (`scaling/opc-sweep.tsv`)

The same shapes through the real managed publication, where the session is
built per package and therefore builds its pool per publication unless the
caller attaches a facility. 40 samples per leg, paired width-1 legs, A/A floors
0.15%-2.07%.

| remainder | w1 | w2 | w4 | w8 | w4 + caller facility | w8 + caller facility |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 7,168 B | 1.000 | 1.001 | 0.997 | 1.015 | 0.998 | 0.987 |
| 28,672 B | 1.000 | 1.016 | **0.933** | **0.848** | 1.281 | 1.217 |
| 61,440 B | 1.000 | 1.167 | 1.187 | 1.142 | 1.411 | 1.382 |
| 114,688 B | 1.000 | 1.201 | 1.294 | 1.222 | 1.439 | 1.379 |
| 245,760 B | 1.000 | 1.240 | 1.366 | 1.306 | 1.432 | 1.482 |
| 458,752 B | 1.000 | 1.254 | 1.412 | 1.409 | 1.418 | 1.453 |
| 638,976 B | 1.000 | 1.218 | 1.390 | 1.405 | 1.383 | 1.476 |
| 983,040 B | 1.000 | 1.250 | 1.456 | 1.483 | 1.445 | 1.498 |
| 1,835,008 B | 1.000 | 1.287 | 1.488 | 1.465 | 1.367 | 1.365 |

The 7,168-byte row is below both thresholds and is not admitted at any width:
0.997-1.015, inside its 2.07% floor — the cost of *considering* a wave is
nothing. The 28,672-byte row is the one that made the second threshold
necessary: on a cold pool it loses 6.7% and 15.2% at widths 4 and 8, while on a
caller's facility the same set gains 28.1% and 21.7%. Pool construction costs
about 120 µs, and a publication owns its session.
`DEFAULT_MIN_POOL_PARALLELIZABLE_BYTES` is therefore **64 KiB**, above the
losing row and below the first row that wins cold.

### 4. Real fixtures, end to end (`timing/fixture-timing.tsv`)

Open plus publish of the first sixty-four Deflate-compressed XML parts of a real
fixture, each replaced with a compact re-serialization of its own content plus a
marker comment. 40 samples per leg, 8 warmups, paired width-1 legs, CPU set
`3,4,5,9,13,15,17,24` measured at 0.0-1.0% utilization immediately before the
run. A/A floors 0.04%-0.83%.

| fixture | tasks | remainder | w1 p50 | w2 | w4 | w8 | w4 + facility | w8 + facility |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `no_drawing_patriarch.xlsx` | 8 | 3.40 MB | 84.51 ms | **0.976** | 1.212 | 1.214 | 1.202 | 1.215 |
| `slide-section-test.pptx` | 58 | 644 KB | 7.83 ms | 1.163 | 1.292 | 1.231 | 1.290 | **1.404** |
| `ConditionalFormattingSamples.xlsx` | 59 | 435 KB | 7.09 ms | 1.168 | 1.339 | 1.372 | 1.370 | **1.521** |
| `layout-in-cell-2.docx` | 14 | 282 KB | 4.13 ms | 1.056 | 1.321 | 1.291 | 1.384 | **1.393** |
| `shared_formulas.xlsx` | 8 | 14.8 KB | 388 µs | 1.005 | 0.997 | 0.999 | **1.412** | 1.355 |
| `sortconditionref.xlsx` | 7 | 7.9 KB | 272 µs | 1.006 | 1.010 | 1.006 | 1.002 | 1.000 |

The two threshold rows behave exactly as designed: `shared_formulas`, whose
14.8 KB remainder sits between the thresholds, stays serial on a cold session
(0.997-1.005, inside its 0.55% floor) and gains 1.41× on a caller's facility;
`sortconditionref`, below both, is serial everywhere at 1.000-1.010.

**One measured regression.** `no_drawing_patriarch.xlsx` at width 2 is 0.976×,
a 2.4% loss against a 0.35% floor. Its changed set is one 3.44 MB member and
seven smaller ones, and at width 2 the wave's two halves are the pole plus the
next-largest members against the rest — the least favourable split of a
one-dominant-member set. The same set gains 1.21× at widths 4 and 8, where the
stealing has somewhere to go. The loss is reported rather than averaged away;
it is below the 5% review threshold and no scenario exceeded it.

**How much of a save this is.** The open is 0.05%-6.6% of these timings
(`open-only` rows, 29-472 µs against 272 µs-84.5 ms), so these figures are
publication-dominated. The end-to-end gain of 1.21-1.52× is well below change
0624's modelled 1.85×, and the reason is visible in the same table: this route
audits every overlay payload as authored XML and re-reads and copies every
unchanged member, and neither is compression. 0624's modelled figure was
Amdahl arithmetic over an *authored* save's deflate share; this is a
source-backed publication measured end to end.

### 5. Allocation and page faults (`counts/faults.txt`)

200 publications of `ConditionalFormattingSamples.xlsx`, `/usr/bin/time`:

| leg | minor faults per publication | peak RSS |
| --- | ---: | ---: |
| width 1 | 292.6 | 7.3 MiB |
| width 4 | 12.7 | 9.4 MiB |
| width 8 | 18.1 | 12.3 MiB |
| width 4, caller facility | 12.0 | 9.2 MiB |
| width 8, caller facility | 16.8 | 11.2 MiB |

Change 0618's gate was that per-worker compressor state must not reproduce the
271-minor-fault pathology it measured. It does not arise: this change keeps one
compressor per member, and the measured faults per publication *fall* by 23×,
because several threads' allocator arenas retain the freed compressor buffers
that one thread's kept returning to the system. Peak RSS rises 2.1 MiB at width
4 and 5.0 MiB at width 8 — 525 KiB and 625 KiB per worker — which is why
`WORKER_STATE_BYTES`, the figure a wave reserves against `Resource::Memory`, is
640 KiB rather than zlib's bare 256 KiB footprint.

### 6. The corpus's changed sets (`counts/changed-set-census.tsv`)

316 of 321 OOXML fixtures under `test-data/` yield an eligible changed set under
this scenario: 1 to 59 tasks (median 10), total payload 1.1 KB to 6.8 MB
(median 28.5 KB), remainder 0 to 3.4 MB (median 15.4 KB). 226 of 316 (71.5%)
clear the 8 KiB warm threshold with two or more tasks; 85 (26.9%) clear the
64 KiB cold one.

## Correctness evidence

**Byte identity over the corpus, at every width** (`identity/corpus-identity.txt`).
Every OOXML fixture under `test-data/` (321 files), published once unmanaged
(the sequential path) and then at widths 1, 2, 4 and 8 and once more through a
caller-supplied Rayon facility, digests compared:

* **305 identical** — the same SHA-256 at every width and on the facility;
* **11 refused identically** — a signature policy, a source with trailing
  bytes, a non-canonical UTF-N member or a malformed XML declaration refuses
  the publication before any width matters, and every managed leg reproduces
  the same error text;
* 5 skipped by the probe before any publication (no eligible part, or a
  catalog the probe's scenario cannot address).

**Determinism under repetition** (change 0625's method): a test publishes the
same plan seventeen times through one width-8 session and asserts every output
equal to the first, with the pool live throughout.

**Tests added.** In `soapberry-zip`: identical bytes at widths 1/2/4/8 with the
pool count asserted per width; determinism across seventeen repeats; a caller
facility producing the sequential bytes with zero pool threads and exactly one
wave of twelve tasks; the balance rule keeping a 4 MiB-plus-631-byte set
serial; the rule admitting forty 871-byte members and narrowing by an explicit
per-task floor; a cold session admitting only from the larger remainder and
becoming warm once its pool exists; stored and precompressed members never
being tasks; the sequential error returned in plan order with the sink
untouched; a cancelled wave publishing nothing; physical local order preserved
at width 4 when it differs from central order; an appended member scheduled
with the regenerated ones. In `litchi-opc`: byte identity at widths 1/2/4/8
with the worker permits observed at the first output byte (0 at width 1, W
above it) and `CpuTasks` charged once per task at every width; a budget that
grants no worker permits narrowing to the sequential bytes instead of failing;
a four-task `CpuTasks` budget refusing fifteen tasks with
`ResourceLimit { resource: CpuTasks }` and an untouched sink; a cancelled
managed publication returning `OpcError::Cancelled` with an untouched sink; a
caller facility running every task with the budget charged and no permits
retained.

**Boundary audits.** No ordinary CRUD signature names `ScopedWorkers`,
`ParallelWriteSession` or `DeflateWave`: they occur only in
`litchi-core/src/execution.rs`, `soapberry-zip`'s `office` and `preserve`
modules, and `litchi-opc/src/source_backed.rs` (private). `soapberry-zip` still
does not depend on `litchi-core` — it defines its own `ScopedWorkers` and
`litchi-opc` bridges them. No `unsafe` is added; `#![forbid(unsafe_code)]`
holds where it held.

**Gates run** (`gates.txt`): `cargo fmt --all --check`; `cargo clippy -p
litchi-core -p soapberry-zip -p litchi-opc --all-targets`; `cargo test -p
litchi-core -p soapberry-zip -p litchi-opc`; `cargo doc` for the three crates;
the consumer suites (`litchi-xlsx`, `litchi-docx`, `litchi-pptx`, `litchi-xlsb`,
`litchi-xls`, `litchi-odf-common`, `litchi-odt`, `litchi-ods`, `litchi-odp`,
`xml-minifier`); `cargo test -p litchi --features docx,xlsx,pptx,xls`;
`python3 tools/non_iwork_gate.py verify`; and the harness's own suite. The
pre-existing failures the briefing lists reproduce unchanged. One further
pre-existing failure is recorded rather than fixed: `cargo test --workspace`
fails to compile `litchi-iwa`'s `create_pages_pie_chart` example, a type
mismatch between `litchi_iwa::Error` and `litchi_iwa_common::chart::data::DataError`
in an iWork crate this change does not touch and the briefing excludes.

## Validation preserved

No limit, refusal, audit or fence moved. The compressed-size budget is still
evaluated per write call; the mini-archive structural checks still run inside
`generated_entry` for every member, whichever thread compresses it; the overlay
and source XML audits, the signature policy, the duplicate-Part checks and the
publication plan are untouched; `finish_source_publication` remains
authoritative for `SourceChanged` and `IncompleteOutput`. The new refusals are
additive and typed: a `CpuTasks` or `Workers` charge that a caller's own budget
cannot accept, refused before any member is compressed, and a worker-pool
construction failure. A task that panics is turned into
`ParallelWriteWorkerPanic { ordinal }` rather than unwinding through a caller's
executor, which is why the facility's contract can say tasks never unwind.

## Limitations — what is not claimed

* **No claim is registered**, and no saving is claimed for the documented
  ordinary save, the eager `PackageWriter` route, any unmanaged package, or the
  splice/replay route. Those routes carry no execution context, cannot take a
  wave, and measured unchanged.
* The streaming writer is out of scope, as change 0624 §4.2 set out, so the
  authored path — where 0624's largest modelled figure lived — is untouched.
* Every figure is scoped to this host, this build, these fixtures and these
  scenarios, and to a machine with free CPUs. **On a saturated host the wave is
  a large loss**: an earlier run of the same sweep whose pinned set happened to
  include three fully-busy CPUs measured 0.005×-0.3× on shapes that measure
  1.2×-7.4× on idle CPUs, because a wave must wait for its workers to be
  scheduled. Neither threshold protects against that; only the caller's
  decision to opt in, and the `Resource::Workers` ceiling a coordinator can set
  across sessions, do. The contended run is retained in
  `scaling/contended-window.tsv` as the counter-example.
* The changed sets are the probe's scenario — the first sixty-four Deflate XML
  parts of a package, re-serialized compactly with a marker comment — not the
  changed sets a particular editor produces. They are realistic in size and
  member count (they come from the fixtures) but their *content* is a compact
  re-serialization, and deflate's rate varies about 4× with content (0624).
  Real packages are usually not compact (94 of 95, change 0602), so this
  scenario is what an editor's own compact writer would emit, not what the
  producer wrote.
* The width-2 regression on a one-dominant-member set (§"Measured" 4) is
  reported, not diagnosed: the split hypothesis is untested.
* `Resource::IoConcurrency`, ADR 0031 §1's third dimension, is **not**
  implemented here. It bounds concurrent positional reads, which belongs to the
  three read sessions ADR 0031 §6 names; adding a dimension no session consumes
  would be an unenforced bound. Gates G1 (the lazy read pool) and G3 (three
  sessions sharing one root) of change 0615 likewise remain open: this change
  makes the *write* session lazy and budget-bound, and leaves the read sessions
  as they are.
* Cold-cache, physical-device, cross-platform and compression-level behaviour
  were not measured. Compression level remains fixed at 6.
* No instruction counts were taken. Callgrind mis-prices exactly what this
  change moves (bulk `rep movsb` per byte, change 0604) and cannot observe
  concurrency at all; cycles and wall time are the metrics, with the A/A floor
  beside them.

## Retained evidence

[`results/change-0662/README.md`](results/change-0662/README.md) — the probe and
its driver scripts, the corpus identity oracle, the changed-set census, the two
balance-rule sweeps and the contended counter-example, the OPC-route crossover,
the real-fixture timings, the page-fault counts, the harness A/B, `gates.txt`,
`decision.json` and `log-sections.md`.
