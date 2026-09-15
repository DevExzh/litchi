# 0624: deflate is 24-62% of every multi-member save, and the 256 KiB threshold is wrong in both directions

Status: design, retained. `performance_claim: none`. **No production code
changed.** This record measures the gate that item CORE-4 / SAVE-6 was waiting
on, freezes the design the gate justifies, and states the prerequisite that
stops it being built today.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This record answers item **CORE-4 / SAVE-6** of
[0587](0587-remaining-opportunity-survey.md) (rank 36), written up from the two
sides in §"Core, facade, execution and parallelism" (CORE-4) and §"OOXML save
and publication" (SAVE-6). Base `aba3f943d` (change 0618); branch
`perf/0624-parallel-changed-member-deflate-design`.

## Why this record exists

CORE-4 says every writer deflates changed members serially, and sets a gate:

> The retained scaling evidence (0009: OPC 4.52× at 12 workers with a serial
> fraction of about 15%; many-small tasks 0.73× and 0.52×) sets the gate at two
> or more changed members each at or above `min_parallel_bytes` […]. Falsified
> if the median multi-part edit changes fewer than two members above the
> threshold, **or deflate of the changed set is under 20% of save wall time**.

SAVE-6 states the threshold concretely: "pays only at two or more regenerated
members of 256 KiB or more […] so one-cell and one-paragraph saves gain
nothing." Neither half had ever been measured; both halves are measured here,
and the result is the opposite of what the item assumed on both counts.

1. **Deflate is not a minor term. It is the largest single term in a save**, at
   **24.19% to 61.94%** of native user cycles across five scenarios — real
   fixtures, harness corpora and the authored path alike. The 20% half of the
   gate is met everywhere, with the smallest margin 1.2× over the bar.
2. **No real fixture clears the 256 KiB half**, and it would not matter if one
   did. Over the 336 OOXML fixtures under `test-data/`, across every save four
   realistic edit routes could produce, **zero** regenerate two or more members
   of 256 KiB or more. Exactly one fixture
   regenerates even one such member — and that fixture is the one case measured
   here that parallel deflate **cannot** speed up at all (0.997× at every width),
   because its changed set is one 3.38 MB member and one 631-byte member.
3. **The threshold is wrong in both directions, and the right one is about
   balance, not size.** A real save that the 256 KiB rule excludes — 40
   regenerated `.rels` members averaging 871 bytes — reaches **3.67× at width
   4**. A closed-form bound over the changed set's size distribution predicts
   every measured cell to within 8%, and it is the rule the design uses.
4. **It still should not be built yet.** Not because of the numbers, but because
   `docs/GOAL.md` rule 9 requires CPU parallelism to be controlled by an
   explicit execution context, the write path has none,
   [proposed ADR 0031](../adr/0031-execution-context-budgets.md) that would give
   it one is not accepted, and change [0615](0615-execution-context-completeness-design.md)
   §7 makes acceptance and gates G1-G3 the stated prerequisite for exactly this
   work. §"Why this is designed and not built" states that in full.

## Method and provenance

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; `rustc 1.95.0`;
`perf` 7.0.14. Seven other agents were active on the host throughout; run-window
load average 7.7-20.5. Two scratch probes, both retained in the packet, both
built `--release` with `lto = true`:

* `probe0624` (`litchi-opc` + `litchi-xlsx` + `litchi-pptx` + `litchi-sheet`
  path dependencies) drives the real editor and publication routes, censuses
  the changed member set, and times a save against the deflate of that set.
  sha256 `69516155c58b2f9dc04bcad222fd808eaec890298f32262e243b1ec322f37244`.
* `probe0624-deflate` (flate2 + rayon only, **no litchi dependency**) measures
  the deflate scaling curve. Its `zlib-rs` is pinned to the workspace lockfile's
  0.6.7 with `cargo update -p zlib-rs --precise`, so its numbers are comparable
  with a save profile. sha256
  `dc35e7b44bb6ff5fd1d82955044ad9719b0372488776fc82c9ad98257f23380a`.

Counts first, timing last. Deterministic counts were taken with `taskset -c 18`;
scaling legs with `taskset -c 18-25`, so a width-8 leg has exactly eight CPUs
and contends with the other agents for them. Every scaling table carries its own
A/A floor, measured in the same window.

**How a "changed member" is identified.** The probe publishes, then compares the
published archive with the source member by member and calls a member *changed*
when its compressed payload, CRC or method differs. On the preservation path
that set is exactly `PreservationAction::Regenerate`, because `Copy` members are
raw byte copies of the source span (`crates/soapberry-zip/src/preserve.rs:880`,
`:897`). On the authored path there is no source, and change 0607 established
that every member is regenerated on every save; the probe reports all of them.

### The A/A floor

Four repeats of each leg, same window, same pinning (`scaling/aa-floor.txt`):

| set | serial p50 spread over 4 repeats | width-2 p50 spread |
| --- | ---: | ---: |
| 2 × 1,893,450 B | **1.04%** | 0.45% |
| 40 × ~871 B | **0.57%** | 3.62% |
| 4 × 26,698 B | **0.21%** | 1.71% |

These are pure-CPU loops with no I/O and no source observation, so the floor is
far below this host's usual p50 4% / p99 14%. Every speedup reported below is
one to two orders of magnitude above it. The one figure that is **not** above
its floor is the width-1 cost, which is discussed where it appears.

## 1. Measured: the shape of the changed set

### 1.1 The corpus cannot supply the members the threshold asks for

Before any edit: of the 336 OOXML fixtures under `test-data/`, **8 contain a
single member of 256 KiB or more and 5 contain two or more**, and in four of
those five the large members are embedded fonts or media that no text or cell
edit regenerates. The corpus's member-count and byte-volume poles are
anti-correlated: the 132-member `ConditionalFormattingSamples.xlsx` has a
59,242-byte largest member, and the 6.8 MB `no_drawing_patriarch.xlsx` has 13
members.

### 1.2 What a real edit regenerates

`census/census-opc.tsv` (325 fixtures × 2 scenarios) and
`census/census-xlsx.tsv` (157 workbooks × 2 scenarios), one line per fixture:

| route | scenario | n | changed members (min/median/max) | largest changed member, corpus-wide | fixtures with ≥2 changed ≥256 KiB |
| --- | --- | ---: | --- | ---: | ---: |
| `OpcPackage` publish | one external relationship added | 325 | 1 / 1 / 1 | 4,725 B | **0** |
| `OpcPackage` publish | a relationship on every related part | 325 | 1 / 2 / 43 | 6,496 B | **0** |
| `litchi-xlsx` value editor | one cell | 157 | 2 / 2 / 4 | 3,382,600 B | **0** |
| `litchi-xlsx` value editor | one percent of used cells | 153 | 2 / 2 / 4 | 213,058 B | **0** |
| `litchi-pptx` authored save | `Package::new()`, 50 slides | 1 | 161 / 161 / 161 | 11,985 B | **0** |
| `litchi-pptx` authored save | `Package::new()`, 200 slides | 1 | 537 / 537 / 537 | 36,836 B | **0** |

The median realistic edit regenerates **two members totalling about 4 KB**. The
single corpus-wide maximum, `no_drawing_patriarch.xlsx` under a one-cell edit,
regenerates `xl/worksheets/sheet1.xml` at 3,382,600 B **and one 631-byte
member** — one large member, not two.

### 1.3 The harness corpora

`tools/perf-baseline`'s XLSX corpus (`build_xlsx_corpus`, `src/lib.rs:20590`;
`XlsxShape`, `:795-829`) was rebuilt by the probe to the same three shapes and
censused. Its one-percent update set is spread across *every* sheet
(`xlsx_one_percent_updates`, `:21758`), so it regenerates one worksheet member
per sheet:

| shape | sheets × rows × cols | members | one-cell edit changes | one-percent edit changes |
| --- | --- | ---: | --- | --- |
| tiny | 3 × 8 × 8 | 8 | 1 × 1,839 B | 3 × 1,841 B |
| medium | 4 × 32 × 32 | 9 | 1 × 26,228 B | 4 × 26,698 B |
| **dense-wide** | 2 × 256 × 256 | 7 | 1 × 1,860,130 B | **2 × 1,893,450 B** |

`xlsx_one_percent_commit_save` on the dense-wide shape is **the only scenario in
the entire program** that clears SAVE-6's threshold as written: exactly two
regenerated members, each 1.81 MiB. It is a generated corpus, and change 0587's
first finding is that the generated corpora are not the shape real files take,
so a gate that only this scenario passes is a weak pass. §3 shows the threshold
is the wrong test anyway.

## 2. Measured: deflate's share of a save

Two independent methods, and they agree.

**Native cycles.** `perf record -e cycles:u -F 4999`, pinned, over a loop that
publishes an already-mutated package N times, summing every `zlib_rs`, `flate2`
and `deflate` symbol (`timing/perf-*.txt`):

| scenario | changed set | zlib/deflate share of `cycles:u` |
| --- | --- | ---: |
| harness dense-wide, one-percent edit | 2 × 1,893,450 B | **58.85%** |
| `no_drawing_patriarch.xlsx`, one cell | 3,382,600 B + 631 B | **61.00%** |
| authored PPTX, 50 slides | 161 members, 249,666 B | **61.94%** |
| `ConditionalFormattingSamples.xlsx`, relationship on every related part | 40 × ~871 B | **44.18%** |
| `ConditionalFormattingSamples.xlsx`, one cell | 4 members, 19,391 B | **24.19%** |

On the dense-wide save a single symbol,
`zlib_rs::deflate::longest_match::longest_match`, is **45.25%** of the save; the
next is `PackageWriter::validate_authored_xml` at 23.98%.

**Wall-clock ratio.** The same probe times the publish and then times deflating
the same changed set's payloads with the same codec, 200-500 samples per leg
(`timing/savesplit.txt`):

| scenario | publish p50 | deflate p50 | ratio |
| --- | ---: | ---: | ---: |
| harness dense-wide, one-percent edit | 40,833 µs | 27,331 µs | 66.9% |
| harness dense-wide, one cell | 20,189 µs | 13,685 µs | 67.8% |
| harness medium, one-percent edit | 1,317 µs | 879 µs | 66.8% |
| `no_drawing_patriarch.xlsx`, one cell | 39,065 µs | 27,340 µs | 70.0% |
| authored PPTX, 50 slides | 2,563 µs | 1,943 µs | 75.8% |
| authored PPTX, 200 slides | 7,388 µs | 5,660 µs | 76.6% |
| `ConditionalFormattingSamples.xlsx`, all related parts | 576 µs | 353 µs | 61.3% |
| `ConditionalFormattingSamples.xlsx`, one cell | 211 µs | 65 µs | 30.9% |

The ratio method reads high and the native share is the one to quote: the
probe's deflate leg constructs a fresh `flate2` encoder per member, which is
what the preservation writer does today but **not** what the streaming writer
does after change 0618 (one compressor per archive). The two methods bracket the
truth; both clear 20% on every scenario.

**The gate's second half is met with no ambiguity.** The smallest share measured
is 24.19% native, on the smallest save in the set.

## 3. Measured: the scaling curve, and why the threshold is wrong

`probe0624-deflate` runs the writers' own codec — flate2 1.1.10 on `zlib-rs`
0.6.7, `Compression::default()` (level 6), raw deflate, one member per task —
serially and through a caller-owned Rayon pool of a stated width. Nothing in a
real writer can beat these numbers, because a real writer must also frame,
account, audit and emit.

### 3.1 The seven real changed sets

40 samples per leg, 10 warmups, `taskset -c 18-25`
(`scaling/scale-real-sets.txt`, `scaling/scale-authored.txt`). Every set is the
actual member payloads a save regenerated, dumped from the published archive:

| changed set | members | ≥256 KiB | serial p50 | w1 | w2 | w4 | w8 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| harness dense-wide one-percent | 2 | **2** | 27,206 µs | 0.998 | **1.988** | 1.987 | 1.983 |
| `no_drawing_patriarch` one cell | 2 | **1** | 27,360 µs | 0.998 | **0.997** | 0.995 | 0.996 |
| harness medium one-percent | 4 | 0 | 903 µs | 0.986 | 1.959 | **3.859** | 3.870 |
| `ConditionalFormattingSamples` one cell | 4 | 0 | 65 µs | 0.980 | 1.874 | **2.761** | 2.650 |
| `ConditionalFormattingSamples` all rels | 40 | 0 | 353 µs | 0.974 | 1.900 | **3.721** | 4.741 |
| authored PPTX, 50 slides | 161 | 0 | 1,959 µs | 0.997 | 1.860 | **3.873** | 4.703 |
| authored PPTX, 200 slides | 537 | 0 | 5,683 µs | 0.995 | 1.923 | **3.546** | 4.739 |

Read the first two rows together. They are the same total work to within 11%,
and they are the *only* two real-or-harness changed sets in the corpus that
contain a member over 256 KiB. One reaches the arithmetic maximum for two tasks;
the other cannot be helped at any width. SAVE-6's threshold admits both and
distinguishes neither. Meanwhile the 40-member `.rels` set — which SAVE-6
explicitly excludes, and whose members average 871 bytes — is the best-scaling
real set in the table.

### 3.2 There is no lower size threshold in the range that matters

A controlled crossover sweep at a fixed 871-byte member size, varying only the
member count, 80 samples per leg (`scaling/crossover.txt`):

| set | serial p50 | w1 | w2 | w4 |
| --- | ---: | ---: | ---: | ---: |
| 4 × 871 B | 34.5 µs | 0.945 | 1.766 | 2.724 |
| 8 × 871 B | 69.5 µs | 0.975 | 1.889 | 3.162 |
| 16 × 871 B | 139.8 µs | 0.985 | 1.914 | 3.471 |
| 32 × 871 B | 279.3 µs | 0.973 | 1.892 | 3.580 |
| 40 × 871 B (the real `.rels` set) | 354.8 µs | 0.978 | 1.922 | 3.671 |
| 64 × 871 B | 564.5 µs | 0.993 | 1.925 | 3.658 |
| 128 × 871 B | 1,135.5 µs | 1.005 | 1.993 | 3.853 |
| 256 × 871 B | 2,235.3 µs | 0.999 | 1.985 | 3.854 |

Deflate parallelizes from **four members of 871 bytes** upward — a 3.5 KB
changed set, three orders of magnitude under the proposed floor. This is not the
many-small regime changes 0009, 0498 and 0499 measured, and the reason is
structural: those were *decompression batches* contending on one positional
source with per-wave thread creation and per-task memory reservations; a deflate
task borrows an owned slice, touches nothing shared, and returns a `Vec`.

The **width-1 cost is 0.5% to 5.5%** — `pool.install` plus one block-and-wake,
about 3-5 µs per call. It is the price of taking the parallel path and finding
one task, and on sets under about 50 µs it is not distinguishable from the A/A
floor.

A wider synthetic sweep over 7 member sizes × 6 member counts × 4 widths is
retained in `scaling/sweep-xml.txt`. Its width-8 column and its smallest cells
are **not usable**: a second run of the same shapes (`scaling/sweep-small.txt`)
put several of them below 1.0×, and the controlled crossover above reproduces
the same shapes above 1.7× at width 2. Those cells are host contention, are
reported as such, and nothing is claimed from them.

### 3.3 The rule that does predict the measurements

For a changed set of member sizes `s₁ ≥ s₂ ≥ … ≥ sₙ` at width `W`, model each
member's deflate time as `tᵢ = f + sᵢ/r` (a fixed per-member term plus a rate),
and bound the speedup by the longest pole against the ideal split:

```
speedup ≤ Σtᵢ / max( t₁ , Σtᵢ / W )
```

Against the measured cells, with `f ≈ 8 µs` and `r` fitted per set (it is
content-dependent: 1.7 ns/byte on `.rels` and small worksheet XML, 7.2 ns/byte
on the dense integer grids, because match search dominates on the latter):

| set | W | bound | measured | error |
| --- | ---: | ---: | ---: | ---: |
| 2 × 1,893,450 B | 2 | 2.00 | 1.988 | 0.6% |
| 2 × 1,893,450 B | 4 | 2.00 | 1.987 | 0.7% |
| 3,382,600 B + 631 B | 2 | 1.00 | 0.997 | 0.3% |
| 3,382,600 B + 631 B | 8 | 1.00 | 0.996 | 0.4% |
| 4 × 26,698 B | 4 | 4.00 | 3.859 | 3.5% |
| 4 members, 19,391 B | 4 | 2.65 | 2.761 | 4.2% |
| 40 × ~871 B | 4 | 4.00 | 3.671 | 8.2% |

The two authored rows are the same story from the other end: 161 members
averaging 1,550 bytes, none within two orders of magnitude of 256 KiB, reaching
3.873× at width 4 — the best-scaling set measured, and the one the threshold
would have excluded most confidently.

**The predictor is the second-largest member, not the largest.** A changed set
is worth splitting when `Σtᵢ − t₁` is large enough to pay for the pool; it is
worthless when one member is the whole set, whatever that member's size.

### 3.4 What this is worth end to end (modelled, not measured)

Amdahl over the measured section share (§2) and the measured section speedup
(§3.1). These are **modelled** figures: no end-to-end parallel save exists to
time, and none is claimed.

| save | deflate share | best measured section speedup | modelled save speedup |
| --- | ---: | ---: | ---: |
| authored PPTX, 50 slides | 61.94% | 3.873× at w4 (161 members) | **1.85×** |
| `ConditionalFormattingSamples`, all rels | 44.18% | 3.67× at w4 | **1.47×** |
| harness dense-wide, one percent | 58.85% | 1.99× at w2 | **1.41×** |
| `ConditionalFormattingSamples`, one cell | 24.19% | 2.76× at w4 | **1.18×** |
| `no_drawing_patriarch`, one cell | 61.00% | 1.00× | **1.00×** |

The largest modelled win is on the **authored** path, which is also the harder
of the two writer boundaries (§4.2).

## 4. The frozen design

**Name.** Per-member parallel deflate of a save's changed set, under an explicit
execution context.

### 4.1 The boundary: `PreservationIndex::prepare`, and nothing else first

`crates/soapberry-zip/src/preserve.rs:840-914`. `prepare` walks the plan's
actions, and for each `PreservationAction::Regenerate` calls
`generated_entry(entry)` (`:897`, and `:911` for appended members), writing the
result into `prepared[index]`. `generated_entry` (`:2070`) is a **pure function
of one `&RegeneratedEntry`**: it compresses the payload into a one-entry
`ZipArchiveWriter`, re-parses that mini-archive to locate its framing, and
returns a `PreparedEntry`. It touches no shared state, borrows nothing from the
index, and allocates only its own buffers.

Four properties make this boundary nearly free of design risk, and all four are
properties the code already has:

* **Output order is not the loop order.** Results are placed at `prepared[index]`
  and emitted later by `self.local_order` (physical order, `:780`) and by index
  (central-directory order, `:808`). The plan's action list "is not an ordering
  control" is the writer's own documented contract (`:236-243`). Reordering
  *when* a member is compressed therefore cannot move a byte.
* **Every compression already completes before any byte is emitted.**
  `write_to_with_accounting` (`:767`) calls `prepare` first and only then writes
  local records. Parallelism does not move the emission boundary.
* **Peak memory does not rise.** `prepare` already holds every regenerated
  member's mini-archive buffer simultaneously — the fact change 0618's rejected
  pooled-compressor variant turned on. A width-`W` version holds the same set
  plus `W` compressor states of about 300 KiB. That is the only new resource,
  and it is bounded by `W`, not by the member count.
* **Deflate reuse survives.** Change 0618 made the preservation writer construct
  one `ReusableDeflateState` per member deliberately, because a compressor
  carried across members is freed under the retained buffers and costs 271 minor
  faults per publish. Per-*worker* reuse has the opposite allocation shape (the
  state is thread-local and outlives nothing), so the design keeps one state per
  worker and must re-measure that against 0618's finding rather than assume it.

### 4.2 The streaming writer is explicitly out of the first implementation

`StreamingArchiveWriter` (`crates/soapberry-zip/src/office.rs:5744`) writes each
member into the sink as it compresses it (`write_deflated_with_accounting`,
`:6919`). Parallelising it means compressing ahead of the write cursor into a
reorder buffer, which introduces in-flight bytes that do not exist today, and
`LimitedEntryWriter` evaluates the compressed-size budget *per write call*
(change 0618), so buffering changes when a limit is refused. That is a contract
change under this program's rules, and it is where the authored path's modelled
1.85× lives. It is designed here only to the extent of naming the constraints:
a bounded reorder window of `min(W, max_in_flight_tasks)` members capped by a
new `max_in_flight_output_bytes`, refusal evaluated on the same per-write
granularity by charging the limit inside the worker before the member is queued,
and the write cursor advancing strictly in submission order. **Not in scope for
a first implementation.**

### 4.3 How the execution context governs it

Exactly the shape [ADR 0031](../adr/0031-execution-context-budgets.md) §5-6
prescribes, and no new dependency edge:

* `soapberry-zip` does not depend on `litchi-core` and must not start to. It
  defines its own `ScopedWorkers` trait, structurally identical to
  `litchi-core`'s, exactly as it already defines its own `CancellationProbe`
  (`office.rs:232-249`) for the same reason.
* A new `ParallelWriteSession` beside `ParallelReadSession`, taking a
  `ParallelWriteLimits` of `workers`, `max_in_flight_tasks` and a **per-task**
  floor. It builds its pool **lazily**, on the first batch that qualifies — 0615
  gate G1's requirement, which `ParallelReadSession` does not meet today.
* `litchi_opc::OpenSession` bridges `ExecutionLimits` to it, exactly as it
  already bridges to `ParallelReadLimits`
  (`crates/litchi-opc/src/execution.rs:33-50`), reserving `Resource::Workers`
  when the pool is built and consuming `Resource::CpuTasks` once per regenerated
  member.
* Cancellation is probed before the batch and between waves, never inside a
  member: a task that has started must finish, because `docs/GOAL.md` rule 3
  forbids abandoning started work for a partial result.
* A new opt-in entry point carries the session — `PackageWriter::to_bytes` and
  `save` keep their signatures and their serial behaviour. Nothing an existing
  caller does changes.

### 4.4 The threshold rule, with its measured basis

Admit the parallel path only when **all** hold, computed from the plan before
any compression:

1. at least two members carry `RegeneratedPayload::Owned` or `Shared` with
   `CompressionMethod::Deflate` (`Precompressed` and `Store` members take the
   bounded direct path in `generated_entry` and are never tasks);
2. `Σsᵢ − s₁ ≥ min_parallelizable_bytes`, the **remainder after the largest
   member** — §3.3's predictor. The measured basis for a starting value is
   §3.2: four 871-byte members (a 2.6 KB remainder) already return 2.72× at
   width 4, and the cost of admitting a set that turns out not to split is the
   0.5-5.5% width-1 cost of §3.2. A value in the low tens of kilobytes is
   conservative against both; the implementation reports the value it chose and
   the sweep that chose it.
3. the granted width is `min(W, n, 1 + (Σsᵢ − s₁)/min_task_bytes)`, so a set of
   one dominant member plus crumbs is narrowed to width 1 and takes the serial
   path with no pool built at all. `min_task_bytes` is ADR 0031 §7's proposed
   per-task floor, and this record is a second, independent measurement in its
   favour — with the correction that for *deflate* the right floor is on the
   order of a kilobyte, not the 256 KiB SAVE-6 assumed, and that the aggregate
   `min_parallel_bytes` alone cannot express rule 2 at all.

### 4.5 Admission gates for an implementation

In addition to 0615 §7's G1-G3 and G7, which remain prerequisites:

* **A1 — byte identity over the corpus.** All 336 OOXML fixtures × the five
  scenarios censused in §1.2, at widths 1, 2, 4 and 8, byte-identical to the
  serial leg, with every typed refusal reproduced with identical text. The
  existing oracles in `results/change-0618/oracle/` are the harness.
* **A2 — error identity and error order.** The serial loop returns the *first*
  error in plan-action order. A parallel version must collect every result and
  return the same one — not the first to fail in time. A plan with two members
  that both fail must produce the serial leg's error, at both widths, and a
  refused member must leave every other member byte-identical (the property
  change 0618's `a_refused_sized_member_leaves_the_next_member_byte_identical`
  pins for the serial writer).
* **A3 — no regression at width 1 or below threshold.** At width 1, and for
  every changed set the rule keeps serial, no p50, p95 or p99 outside the A/A
  floor of §"The A/A floor", and no pool constructed (counted from
  `/proc/self/task`).
* **A4 — scaling stated in full.** Widths 1/2/4/8/N with speedup, efficiency and
  Amdahl serial fraction per cell, superlinear and `S < 1` cells labelled
  out-of-model and never fitted, in the shape change 0088's harness emits.
  Efficiency below 25% at the widest width is a review trigger.
* **A5 — allocation and page faults.** Change 0618 rejected a pooled compressor
  because it cost 271 minor faults per publish. Per-worker state must be
  measured against that: minor faults per publish at each width, and peak RSS,
  against the serial leg.
* **A6 — the ordering contract restated as a test.** A fixture whose physical
  local order differs from its central-directory order, published at width 4,
  byte-identical — beside `preserves_physical_local_order_when_central_order_differs`
  (`preserve.rs:3662`).

**Falsification.** The design is falsified if A2 cannot be met without
serialising the error path, or if A5 shows per-worker compressor state
reproducing 0618's allocator pathology at any width.

## Why this is designed and not built

Four reasons, in the order they bind.

**1. The prerequisite is an unaccepted ADR, not a measurement.** `docs/GOAL.md`
rule 9: "CPU parallelism must be opt-in and controlled by an explicit execution
context with thread, memory, I/O, cancellation, and task-granularity budgets."
The entire write path — `litchi-opc/src/pkgwriter.rs`, `atomic.rs`, the
`soapberry-zip` writers, the `litchi-cfb` writers — has **zero**
`ExecutionContext` or `ExecutionLimits` references, which this record re-verified
by grep at the base commit. Giving it one is proposed ADR 0031, which is not
accepted, and change 0615 §7 states plainly: "No part of this design may be
implemented before a human accepts ADR 0031", with G1 (lazy pool), G2 (the
opt-in default costs nothing) and G3 (the budgets compose) ordered before any
*use* of the parallelism. This record is that use. Building it first would
either put an ungoverned pool in the writer — the thing rule 9 names — or
duplicate a budget design a human has not yet reviewed.

**2. The optimization order puts a step-1 item ahead of it in the same
profile.** `docs/GOAL.md` orders elimination of unnecessary work before bounded
parallelism (step 5). The dense-wide save's second-largest symbol is
`PackageWriter::validate_authored_xml` at **23.98%** of the same save, and the
`ConditionalFormattingSamples` one-cell save spends 24.19% on deflate against a
serial remainder still carrying the audit. Parallel deflate divides that
remainder's cost by nothing.

**3. The scenario that clears the threshold as written is a generated corpus.**
§1.2 is unambiguous: 0 of 336 real fixtures regenerate two members of 256 KiB or
more under any of four realistic edit routes. Under the *measured* rule of §4.4
a great many real saves qualify — but that rule is this record's proposal, not
an accepted one, and the honest reading of the gate as the survey wrote it is
that its size half fails on every real file.

**4. The one real fixture with a large regenerated member is the one that cannot
win.** `no_drawing_patriarch.xlsx` is 0.997× at every width. A change whose
headline scenario would be the corpus's largest save, and which does nothing
there, needs its scope stated before it is written, not after.

What would unblock it, in order: ADR 0031 accepted (or a narrower amendment
covering `Resource::Workers`, `Resource::CpuTasks` and a per-task floor); 0615's
G1 landed; 0615's G3 met; then this record's A1-A6.

## ADR compliance

| Record | Reading |
| --- | --- |
| [ADR 0002](../adr/0002-crate-topology.md) | `soapberry-zip` gains no dependency: it defines its own `ScopedWorkers`, as it already defines its own `CancellationProbe`. `litchi-opc` bridges, as it already bridges `ExecutionLimits` to `ParallelReadLimits`. |
| [ADR 0005](../adr/0005-io-memory-and-performance.md) | "CPU parallelism is opt-in through an execution context controlling scheduling, affinity, cancellation, thread and memory budgets." The design adds no path that is not opt-in, and peak memory rises only by `W` compressor states — the preservation writer already retains every prepared member. |
| [ADR 0006](../adr/0006-validation-security-and-compatibility.md) | Published bytes must not change. The design's whole claim to soundness is that `generated_entry` is a pure function and output order is set by `local_order` and index, not by the loop; gate A1 measures it rather than arguing it. |
| [ADR 0011](../adr/0011-ooxml-physical-package-ownership.md) | No archive type, lock or executor becomes visible. `ParallelWriteSession` is a `soapberry-zip` type reached only through an advanced-ingress session parameter, never from a `Workbook`, `Document` or `Presentation` signature. |
| [ADR 0031](../adr/0031-execution-context-budgets.md) (**proposed, not accepted**) | This record is the first measured consumer 0615 predicted, and it supplies a second independent argument for §7's per-task floor — with the correction that a *deflate* floor is kilobytes, not the 256 KiB SAVE-6 assumed, and that the aggregate `min_parallel_bytes` cannot express the largest-member rule at all. |
| `docs/GOAL.md` rules 8, 9, 10 | No ambient Rayon, no global pool, no `unsafe`. The pool is session-owned and lazily built. |

## Validation preserved

No production file was modified, so every limit, refusal, audit and fence is
exactly as change 0618 left it. The design changes none of them: the compressed
-size budget stays per write call inside `LimitedEntryWriter`, the mini-archive
structural checks stay in `generated_entry`, `verify_authored` and the
publication plan are untouched, and §4.5 gate A2 makes error identity and error
*order* an admission requirement rather than an assumption.

## Limitations — what is not claimed

* **No speedup is claimed.** Every §3 figure is deflate measured *in isolation*
  by a probe with no litchi dependency; §3.4's end-to-end figures are Amdahl
  arithmetic and are labelled modelled. No parallel save exists to time.
* No claim is registered. Every number is scoped to this host, this build, these
  fixtures and these scenarios.
* The wider synthetic sweep's width-8 column and its smallest cells did not
  reproduce across windows and are reported as host contention; nothing is
  claimed from them. The width-8 legs everywhere ran on eight CPUs shared with
  seven other agents.
* The deflate-share ratios in §2's second table read high because the probe
  constructs one encoder per member. The native cycle shares are the figures to
  quote and the two bracket the truth; the gap between them is the per-member
  encoder cost change 0618 removed from the streaming writer only.
* `r`, the per-byte deflate rate, varies 4× with content (1.7 to 7.2 ns/byte
  across the sets measured). §3.3's model is fitted per set, not predictive from
  size alone; §4.4 rule 2 therefore uses size as a proxy and relies on the
  bounded cost of guessing wrong.
* The DOCX multi-part case CORE-4 names — headers, footers, footnotes — was not
  measured through `litchi-docx`'s own editor; the DOCX evidence here is the
  `OpcPackage` publication route over all 62 DOCX fixtures. The largest DOCX
  fixture's twelve ≥256 KiB members are embedded fonts, which no paragraph edit
  regenerates.
* Cold-cache, physical-device, peak-RSS, allocation-profile, cross-platform and
  compression-level behaviour were not measured. Compression level remains
  fixed at 6 everywhere and is still unmeasured, as SAVE-6 noted.
* The censuses exclude what each route refuses: 11 of 336 fixtures refused at
  open by the `OpcPackage` route, and — after dropping the 66 DOCX and PPTX rows
  the XLSX route correctly rejects as non-XLSX — 17 one-cell and 21 one-percent
  refusals by the `litchi-xlsx` editor. Every refusal is in the retained TSVs and
  none was investigated here.

## Retained evidence

[`results/change-0624/README.md`](results/change-0624/README.md) — both probes
with their manifests, the corpus censuses, the `perf` symbol reports, the
save-split timings, the real-set and synthetic scaling tables with the A/A
floor and the crossover sweep, `gates.txt`, `decision.json` and
`log-sections.md`.
