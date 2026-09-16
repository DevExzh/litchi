# 0636: a bounded window over the CFB stream cursor takes the XLS validation walk from 82,727 positional reads to 60

Status: retained, implemented in `litchi-cfb` (a new wrapper type) and
`litchi-xls` (one caller). `performance_claim: none` — this record carries
deterministic read, byte and source-observation counts over the whole XLS
corpus, `perf stat` cycles and instructions, callgrind isolation pairs and
paired medians with a measured A/A floor, and registers no claim-registry
entry.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is item 18 of change [0630](0630-queue-refresh-after-the-first-wave.md)'s
queue, priced by change [0627](0627-ole2-range-source-selectors.md). It does not
land where 0630 expected it to, and the reason is the first thing this record
has to say.

## What the brief expected, and what the measurement found

0630 pointed at the whole-sheet walk: 0627 measured `54016.xls` `all-cells` at
**16,145 requests over a range source, 16,061 of them 512 bytes or smaller**, and
named a cursor read-ahead as the lever. The brief's mechanism sentence was that
"the worksheet frame loop asks for small pieces (a 4-byte header, then a
payload), so a range source pays one round trip per record."

An attribution probe over the public `litchi_xls` APIs (retained as
[`probe/main.rs`](results/change-0636/probe/main.rs)) says otherwise. Of the
16,105 positional reads the `54016.xls` walk takes beyond its open:

* the **worksheet scan** takes single digits. Change
  [0568](0568-xls-worksheet-window.md) already gave it a window that grows
  512 → 64 KiB, and the probe sees exactly that schedule
  (`512, 1024, 2048, 4096, 8192, 16384, …`);
* **16,000-odd are shared-string resolutions.** `resolve_shared_string_inner`
  builds a fresh `stream_cursor_at_hinted` cursor per string and reads that
  string's bytes — 6,214 reads of 3 bytes, 3,316 of 11, and so on — jumping
  **backwards 4,444 times**. One read per cursor, in a random order.

A read-ahead in the cursor cannot help that: there is no second read on any of
those cursors to serve, and a window filled on a 3-byte read would be pure
over-read. **The lever 0627 identified is not a cursor read-ahead; it is the
per-string cursor.** That is named for the queue in *Limitations*, and nothing
in this change touches it.

The header-then-payload loop the brief describes does exist, and it is the
**XLS validation** record walk (`StreamingRecordReader` in
`crates/litchi-xls/src/validation.rs`), which reads a four-byte header and then
a payload directly from a cursor, for every record of the Workbook stream. The
same probe prices it:

| fixture | `validate` reads | `all-cells` reads |
| --- | ---: | ---: |
| `54016.xls` | **82,727** | 16,105 |
| `ConditionalFormattingSamples.xls` | **7,388** | 44 |
| `WithCustomViews.xls` | **4,443** | 880 |

`litchi_xls::validation::validate_source` takes an `Arc<dyn ReadAt>`, so it is a
range-source entry point in exactly the sense `docs/GOAL.md` names. On
`54016.xls` it is **five times** the request cost of the walk 0627 called "the
single most expensive thing a range-source caller can ask an XLS reader for",
and 82,693 of its 82,727 reads are 512 bytes or smaller — 41,358 of them are
exactly four.

So this record implements the design the brief asked for, and enables it at the
loop the brief described.

## What was changed

**`crates/litchi-cfb/src/shared.rs`** gains `BufferedOleStreamCursor<'a>`, a
public wrapper that holds one `SharedOleStreamCursor` and a bounded window:

* `new(cursor, ceiling)` clamps `ceiling` to `MAX_READ_AHEAD_BYTES` (64 KiB, the
  bound 0568 gave the worksheet window). Zero leaves the wrapper a passthrough.
* `read_exact` serves from the window when it can; otherwise, for a read no
  larger than the ceiling, it takes **one** `SharedOleStreamCursor::read_exact`
  fill and serves from that; otherwise it drains the window and hands the
  remainder to the cursor.
* Fills are clamped by the ceiling, by the bytes remaining in the stream and by
  what the current read needs, and grow on 0568's schedule — `FIRST_FILL_BYTES`
  = 512, doubling to the ceiling — so the first fill is exact and a caller that
  abandons a stream over-reads at most what it already consumed.
* `skip_forward`/`skip_to` pass resident bytes in memory and hand the rest to
  the cursor, which reads nothing.
* `set_read_ahead` changes the ceiling mid-stream, for a caller that can only
  decide how far ahead it may read after it has read something.

**`crates/litchi-xls/src/validation.rs`**: `StreamingRecordReader` holds a
`BufferedOleStreamCursor` built with ceiling **0**, and `analyze_workbook` calls
`open_window()` — which sets the ceiling to 64 KiB — at the top of the first
loop iteration after `filepass_slot_open` has closed. See *The encryption
boundary*.

Two tests are added to `crates/litchi-xls/tests/validation.rs` and twelve to
`shared.rs`. The tracked diff is +763 / −6 across four files — the wrapper and its
tests in `shared.rs`, one export line in `crates/litchi-cfb/src/lib.rs`, and the
two `litchi-xls` files — plus this record's packet.

### Why a wrapper and not a field on the cursor

The first implementation put the window **inside** `SharedOleStreamCursor`, as
four extra fields, with `read_exact` branching on a ceiling of zero. It was
correct and passed every test, and it cost the read path a measured
**+1.41% cycles and +0.72% instructions on the `54016.xls` whole-sheet walk**,
with paired timing at **+6.20% p50 against a −0.44% A/A floor** — above this
programme's 5% review trigger. The mechanism is that the shared-string resolver
builds one cursor per string: 16,105 cursors on that walk, each now 48 bytes
larger and each carrying `Vec` drop glue where the cursor previously had none.

The wrapper removes that by construction rather than by measurement. Every byte
it moves goes through the **public** cursor API, `SharedOleStreamCursor` is
byte-for-byte the type it was, and a caller that does not build a wrapper cannot
pay for one. The measured read-path delta after the restructure is
**+0.00% instructions on the walk** (callgrind and `perf stat` agree to within
0.00%) and **−0.01% on the open**. Both implementations are in the packet's history; the
field version's numbers are quoted above because they are the reason the
wrapper exists.

## Why it is sound

**No byte outside the selected stream is ever read.** A fill is clamped to
`cursor.len() - cursor.position()`, so the window stops at the declared stream
length exactly as an unbuffered read does. The CFB index, the chain validation,
the `max_input_bytes` and `max_directory_bytes` ceilings and the
`max_workbook_stream_bytes` gate are untouched and run before any fill.

**Bounded memory.** One window of at most the declared ceiling per wrapper,
grown with `try_reserve_exact` and reported as `OleError::Allocation` when the
reservation fails. `MAX_READ_AHEAD_BYTES` is a new ceiling, not a relaxed one.

**Nothing is committed before the step that can fail has succeeded.**
`SharedOleStreamCursor::read_exact` already guarantees that a failed read does
not commit the cursor position; the wrapper truncates its window back to the
pre-fill length on a failed fill and advances `consumed` only after a serve has
succeeded. `a_failed_fill_does_not_commit_the_cursor` retries a failed read on a
stable source and gets the right bytes at the right position.

**Error identity and precedence are the cursor's.** Every fill is a cursor
`read_exact`, so change [0558](0558-ole2-single-read-fence.md)'s trailing fence
and change [0317](changes/0317-opc-source-read-error-precedence.md)'s
`SourceChanged`-wins precedence are executed unmodified. The wrapper adds two
errors of its own: the bounds refusal, which reproduces the cursor's wording
against the caller's logical position, and a corruption report for a window
shorter than the read it serves, which the clamp arithmetic makes unreachable
and which exists so a slice index cannot panic.

**The observation contract moves, deliberately and only for the opted-in
caller.** Change [0621](0621-xls-open-fence-count.md)'s rule is one observation
per read; here **a fill is that read**. A read the window covers publishes bytes
the fill already bracketed and takes no observation of its own, so a mutation
between two served reads is reported by the next fill rather than by the read
after it. Two things bound that:

1. `validate_source_with_limits` ends **every** exit path with
   `shared.source_version()`, which observes the source and refuses
   `SourceChanged`. The operation-level bracket — the version captured at open,
   the version observed after the report is built — is exactly what it was, so a
   mutation anywhere during a validation is still refused with the same typed
   error.
2. A mutation that is reverted before the next fill is not observed. That is
   already the documented model: `litchi_core::FileVersionPolicy` says so for
   reverted transitions, and `SharedOleStreamCursor::read_exact`'s own
   documentation says a mutation reverted before a read's fence is not observed.
   What changes is the width of that window, from one read to one fill.

`read_ahead_reports_a_mutation_at_the_next_fill_not_at_a_served_read` and
`read_ahead_reports_a_mutation_taken_before_its_first_fill` are change 0621's
change-under-read sweep extended to fill boundaries, and they pin both halves,
including that a failed read does not commit.

**ADR reading.** ADR 0003's bounded-resource rule is met by the clamped ceiling
and the fallible reservation. ADR 0005's lazy-payload contract is observed: the
window is filled from the same chain walk, on demand, and never materializes a
stream. ADR 0006 is untouched — no execution context, no worker pool, no ambient
I/O; the wrapper is `!Sync`-by-borrow like the cursor it holds and introduces no
lock. No new `unsafe`, no new dependency, no public leakage of archive types,
raw locks or executors.

### The encryption boundary

`crates/litchi-xls/tests/validation.rs::streaming_validation_stops_before_large_ciphertext_tail`
asserts that validation, on meeting a FILEPASS record, stops **without reading
the ciphertext behind it**. A window opened before that point reads some of it:
with the window on from the first record, that test failed and the corpus sweep
showed `xor-encryption-abc.xls` reading 486 bytes it had never read before.

That is a defence, so the window is not opened until it cannot be crossed.
MS-XLS admits FILEPASS only immediately after the workbook-global BOF, behind
any WRITEPROTECT records; `analyze_workbook` already tracks that as
`filepass_slot_open`. `open_window()` is called at the top of the first loop
iteration after the slot closes — the second or third record — so a
properly-placed FILEPASS is met with the window still shut.

The corpus confirms it: `xor-encryption-abc.xls` reads **8 requests and 2,078
bytes on both legs**, unchanged. A new counting test,
`streaming_validation_opens_no_window_before_a_filepass`, pins it against a
64 KiB ciphertext tail.

The residual is stated rather than hidden: a FILEPASS placed *later* than the
grammar allows is a placement violation that already marks the report
`biff_invalid`, and the walk may have read up to one growth step past it. That
over-read is bounded by what was already consumed and by the 64 KiB ceiling, and
every byte of it is inside the Workbook stream that `max_workbook_stream_bytes`
already bounds.

## Measured

Base commit `c7326f68065edf6f2198ca3cb39c38c48cf00ed9`, branch
`perf/0636-cfb-cursor-range-read-ahead`, both legs built `--release --locked`,
every measured process pinned with `taskset -c 12`, seven other agents active on
the host throughout. Binaries and their SHA-256s are in
[`binaries.sha256`](results/change-0636/binaries.sha256).

### Deterministic counts: every XLS fixture, four read operations and validation

126 `.xls` files under `test-data/`; 113 open, giving 565 operation rows. Every
row's outcome is frozen as a digest — the worksheet names, the cell-value list,
the text, the report, or the verbatim refusal — and **all 565 are identical on
both legs.**

| operation | reads before | reads after | bytes before | bytes after | observations before | observations after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `open` | 1,666 | **1,666** | 3,693,129 | **3,693,129** | 1,567 | **1,567** |
| `list` | 1,666 | **1,666** | 3,693,129 | **3,693,129** | 1,673 | **1,673** |
| `all-cells` | 26,655 | **26,655** | 2,108,161 | **2,108,161** | 27,019 | **27,019** |
| `full-text` | 29,708 | **29,708** | 3,210,676 | **3,210,676** | 79,061 | **79,061** |
| `validate` | 264,622 | **1,842** | 7,511,925 | 7,511,960 | 275,925 | **1,699** |

The four read operations are **identical read for read, byte for byte and
observation for observation** across the whole corpus. That is the control, and
with the wrapper it is a property of the code rather than a measurement: no file
on those paths changed.

Validation: **−99.30% of reads, −99.38% of observations, +35 bytes** — five
ten-thousandths of one percent — across 113 fixtures. The largest per-fixture
over-read anywhere in the corpus is **3 bytes** (`DateFormats.xls`, 5,629 →
5,632), which is a tail fill clamped by the stream end.

Per fixture:

| fixture | file bytes | reads | bytes | observations |
| --- | ---: | --- | --- | --- |
| `54016.xls` | 984,576 | 82,727 → **60** (−99.93%) | 937,418 → **937,418** | 82,705 → **31** |
| `ConditionalFormattingSamples.xls` | 1,402,368 | 7,388 → **78** (−98.94%) | 1,328,049 → **1,328,049** | 7,387 → **38** |
| `WithCustomViews.xls` | 165,888 | 4,443 → **17** (−99.62%) | 153,509 → **153,509** | 4,461 → **19** |
| `xor-encryption-abc.xls` | 4,096 | 8 → **8** | 2,078 → **2,078** | 9 → **9** |

The three that walk their whole stream read **exactly the same bytes**: a
validation frames every record to the end of the substream, and the window is
clamped by that end.

### What that is worth on a range source (modelled)

At 0627's transport — 1 ms of fixed service per physical request, 100 MiB/s, a
64 KiB maximum range — every read on both legs is at most 64 KiB, so physical
requests equal logical reads. This is 0627's model arithmetic, not observed
service:

| fixture | before | after |
| --- | ---: | ---: |
| `54016.xls` | 82,727 requests → **82.74 s** | 60 → **69 ms** |
| `ConditionalFormattingSamples.xls` | 7,388 → **7.40 s** | 78 → **91 ms** |
| `WithCustomViews.xls` | 4,443 → **4.44 s** | 17 → **18 ms** |

### Range-source neutrality of the read path

0627's five XLS range-source selectors, `54016.xls`, same transport, before and
after:

| case | requests | bytes | request-sequence SHA-256 |
| --- | ---: | ---: | --- |
| `xls_range_source_open` | 40 → 40 | 317,171 → 317,171 | identical |
| `xls_range_source_open_list_worksheets` | 40 → 40 | 317,171 → 317,171 | identical |
| `xls_range_source_open_one_cell` | 66 → 66 | 933,004 → 933,004 | identical |
| `xls_range_source_open_all_cells` | 16,145 → 16,145 | 1,256,139 → 1,256,139 | identical |
| `xls_range_source_open_full_text` | 16,145 → 16,145 | 1,256,139 → 1,256,139 | identical |

The digests are over the **complete ordered `(offset, requested, returned)`
sequence**, so this is the strongest available statement that the read path did
not move. The before leg reproduces 0627's published counts exactly.

### Cycles and instructions

`perf stat -r 5` isolation pairs, N=20 and N=120 iterations differenced and
divided by 100, owned in-memory source, CPU 12:

| scenario | cycles before | cycles after | Δ | instructions before | instructions after | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `validate` `54016.xls` | 7,726,862 | 2,021,420 | **−73.84%** | 28,749,586 | 10,553,806 | **−63.29%** |
| `validate` `WithCustomViews.xls` | 447,759 | 153,680 | **−65.68%** | 1,650,157 | 688,786 | **−58.26%** |
| `validate` `ConditionalFormattingSamples.xls` | 958,189 | 543,945 | **−43.23%** | 3,285,021 | 1,778,949 | **−45.85%** |
| `open` `54016.xls` | 807,975 | 810,413 | +0.30% | 3,991,400 | 3,991,087 | −0.01% |
| `all-cells` `54016.xls` | 34,344,352 | 35,081,174 | +2.15% | 150,037,067 | 150,038,391 | **+0.00%** |
| `full-text` `54016.xls` | 80,969,697 | 91,168,888 | +12.60% | 335,220,647 | 333,408,511 | −0.54% |
| `full-text` `WithCustomViews.xls` | 4,653,653 | 4,634,829 | −0.40% | 16,717,314 | 16,666,127 | −0.31% |

**The read-path rows' cycle column is not a measurement of this change.** Three
`perf stat` windows were run over the same pair of binaries; the instruction
deltas repeat to within half a percent and the cycle deltas do not:

| scenario | cycles across three windows | instructions across three windows |
| --- | --- | --- |
| `all-cells` `54016.xls` | +0.48%, +0.68%, +2.15% | +0.00%, +0.00%, +0.00% |
| `full-text` `54016.xls` | −10.82%, −7.44%, +12.60% | −0.01%, −0.18%, −0.54% |
| `full-text` `WithCustomViews.xls` | −8.16%, +1.04%, −0.40% | −0.08%, +0.34%, −0.31% |
| `open` `54016.xls` | −0.01%, −0.27%, +0.30% | +0.01%, −0.01%, −0.01% |

Nothing those paths execute changed, so the cycle column is measuring host state
on a machine with seven other agents on it. The read-path rows are therefore
read from their instruction counts and their paired medians, and the whole
history is retained in
[`cycles/cycles-run-history.txt`](results/change-0636/cycles/cycles-run-history.txt)
rather than reduced to the friendliest window. The `validate` rows are large
enough that this spread does not reach them: they repeat at −73.84% / −74.00% /
−74.21% cycles across the same three windows.

Callgrind isolation pairs (N=2 and N=6, differenced and divided by 4) put one
number differently, and the difference is the point:

| scenario | before Ir/op | after Ir/op | Δ |
| --- | ---: | ---: | ---: |
| `validate` `54016.xls` | 28,798,952 | 12,375,122 | −57.03% |
| `validate` `ConditionalFormattingSamples.xls` | 4,161,600 | 5,265,886 | **+26.54%** |
| `all-cells` `54016.xls` | 151,737,990 | 151,737,968 | −0.00% |
| `full-text` `WithCustomViews.xls` | 16,947,404 | 16,931,303 | −0.10% |

Callgrind says `ConditionalFormattingSamples.xls` validation costs **26.5% more**
instructions; `perf stat` says it costs **45.9% fewer**, and the clock says it
takes 41.9% less time. Callgrind counts `rep movsb` per byte — change
[0604](0604-cfb-append-reads-design.md) measured a 35× overstatement — and the
window copies the 1,328,049-byte stream once more than the unbuffered walk did.
That fixture's records are large (108 payloads of 8,224 bytes), so the extra copy
is most of its instruction budget and the request saving is smallest. It is the
honest worst case for the design, and the metric that prices a bulk copy says it
still wins. The two read-path controls agree to within 0.10% on both
instruments, which is the statement that the walk and the text projection run
the instructions they ran.

### Paired timing

120 samples per leg, order A1 B1 B2 A2, CPU 12, one complete fresh lifecycle per
sample. `owned` is an in-process `ReadAt` over the file's bytes; `file` is
`litchi_core::FileSource`, where every source observation is an `fstat` and
every read a `pread`.

| scenario | A p50 (µs) | B p50 (µs) | Δ p50 | Δ p95 | Δ p99 | A/A floor p50 | B/B p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `validate` file `54016.xls` | 27,962.7 | 472.1 | **−98.31%** | −98.31% | −98.31% | −0.11% | −0.24% |
| `validate` file `ConditionalFormattingSamples.xls` | 2,490.2 | 146.4 | **−94.12%** | −93.85% | −93.56% | +0.07% | +1.02% |
| `validate` owned `54016.xls` | 1,732.6 | 452.3 | **−73.89%** | −73.78% | −73.18% | −0.68% | +0.11% |
| `validate` owned `WithCustomViews.xls` | 99.8 | 33.5 | **−66.45%** | −68.07% | −62.42% | +0.28% | +0.33% |
| `validate` owned `ConditionalFormattingSamples.xls` | 211.4 | 122.9 | **−41.89%** | −40.53% | −40.79% | +0.06% | −1.31% |
| `open` owned `54016.xls` | 180.3 | 178.3 | −1.06% | −1.64% | −2.64% | +0.23% | +0.19% |
| `full-text` owned `WithCustomViews.xls` | 1,029.8 | 1,035.9 | +0.59% | +0.35% | +1.10% | +0.36% | +1.78% |
| `full-text` owned `54016.xls` | 17,650.5 | 17,830.0 | +1.02% | +1.80% | +2.18% | −0.47% | −0.03% |
| `all-cells` owned `54016.xls` | 7,664.1 | 7,903.1 | +3.12% | +2.63% | +2.27% | −1.20% | +1.51% |

**The file-source legs are the headline and the owned legs are the floor.** On
`54016.xls` a validation over `FileSource` falls from 27.96 ms to 0.47 ms
because 82,727 `pread` calls and 82,705 `fstat` calls become 60 and 31.

**The two read-path rows that moved were chased.** In an earlier, busier window
(load average 36–46) `all-cells` read +2.88% and `full-text` on `54016.xls` read
**+6.92%** at p50 — above this programme's 5% review trigger — on paths whose
instruction count is +0.00% and −0.54%. Its two B legs disagreed by 7.48% in
that window, so both scenarios were re-run twice more at 150 samples, and both
were re-run again in the final window:

| scenario | earlier window | repeat 1 | repeat 2 | final window |
| --- | ---: | ---: | ---: | ---: |
| `full-text` owned `54016.xls` | +6.92% (floor +0.39%) | **−4.36%** (floor +11.43%) | **−1.70%** (floor +2.37%) | **+1.02%** (floor −0.47%) |
| `all-cells` owned `54016.xls` | +2.88% (floor +0.91%) | +2.78% (floor −0.95%) | +0.62% (floor +2.13%) | **+3.12%** (floor −1.20%) |

`full-text` crosses zero; `all-cells` sits between +0.5% and +3.2% with its own
A/A floor reaching 2.1% in the other direction. With a **+0.00% instruction
delta** on the walk — the same number from callgrind and from `perf stat`, in
three independent windows — the conclusion is that the read path executes the
same work and the residue is placement and host state. It is reported rather
than averaged away, and every leg of every window is retained in the packet,
including the ones that read worse.

## Correctness evidence

**Twelve new unit tests in `crates/litchi-cfb/src/shared.rs`:**

* `a_cursor_without_read_ahead_reads_and_fences_once_per_call` — the frozen
  baseline: 64 four-byte reads cost 64 requests and 64 observations.
* `a_window_with_no_ceiling_reads_exactly_as_the_bare_cursor_does` — identical
  request list and observation counts, reads and skips.
* `read_ahead_serves_a_framed_walk_from_one_fill_per_run` — 2,048 four-byte
  reads cost **six** requests, and the request list is asserted exactly:
  `(512,4) (516,512) (1028,1024) (2052,2048) (4100,4096) (8196,508)`.
* `read_ahead_never_reads_past_the_declared_stream`,
  `read_ahead_over_reads_at_most_one_growth_step`.
* `read_ahead_reports_a_mutation_at_the_next_fill_not_at_a_served_read`,
  `read_ahead_reports_a_mutation_taken_before_its_first_fill` — 0621's
  change-under-read sweep at fill boundaries.
* `a_failed_fill_does_not_commit_the_cursor`.
* `read_ahead_matches_a_plain_cursor_on_a_fragmented_chain` (a FAT chain whose
  sectors are stored out of logical order),
  `read_ahead_matches_a_plain_cursor_on_a_minifat_stream` (which also asserts the
  Mini Stream cache stays unmaterialized).
* `a_read_larger_than_the_ceiling_bypasses_the_window`,
  `read_ahead_skips_inside_the_window_without_reading`,
  `the_read_ahead_ceiling_is_clamped_to_the_published_bound`.

**Two new integration tests in `crates/litchi-xls/tests/validation.rs`:**
`streaming_validation_frames_many_records_without_a_read_each` (2,000 framed
records cost fewer than 64 reads and no read exceeds the ceiling) and
`streaming_validation_opens_no_window_before_a_filepass`.

**The corpus differential** above is also the oracle: 565 frozen outcomes over
113 fixtures, identical on both legs, covering the worksheet listing, one
worksheet's complete cell-value list, the workbook text projection, the
validation report and every typed refusal — including
`ConditionalFormattingSamples.xls`'s `source XLS parse error: Invalid record
0x0006: shared Formula metadata requires a leading PtgExp token`, which is
reproduced verbatim after 341 reads and 224,946 bytes on both legs.

**Gates**, all in the worktree, tails in
[`results/change-0636/gates.txt`](results/change-0636/gates.txt):

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-cfb -p litchi-xls --all-targets` | clean, 0 warnings (workspace lints are deny) |
| `cargo doc -p litchi-cfb -p litchi-xls --no-deps` | clean (rustdoc lints are deny) |
| `cargo test -p litchi-cfb` | 375 passed, 0 failed, 1 ignored across 5 suites |
| `cargo test -p litchi-xls` | 1,401 passed, 0 failed, 1 ignored across 72 suites |
| `cargo test` for the seven other `litchi-cfb` consumers | 2,701 passed, 0 failed, 24 ignored across 94 suites |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | 266 passed, 0 failed, 7 ignored across 26 suites |
| `cargo test --release --locked` in `tools/perf-baseline` | 531 passed, 0 failed, 1 ignored across 19 suites |

No pre-existing failure was met, so none had to be reproduced on the before
checkout.

## Validation preserved

No validation path was relaxed, bypassed or moved. The CFB index validation, the
allocation-chain checks, the directory and input ceilings, the
`max_workbook_stream_bytes` gate, `max_biff_records`, `MAX_RECORD_BYTES`, the
BIFF grammar checks and the FILEPASS placement rules all run exactly where they
ran, on exactly the bytes they ran on. The one place where validation order and
reading interact — the encryption stop — is the reason the window opens late,
and the corpus proves the encrypted fixture's read set is unchanged to the byte.

## Limitations

**What is not claimed.**

* **No claim-registry entry, and no speedup claim beyond these scenarios.** The
  numbers are scoped to nine timing scenarios, seven cycle scenarios and 113
  fixtures on one host, one toolchain and one build.
* **The range-source figures for validation are modelled, not measured.** They
  are 0627's model arithmetic over a deterministic request count — 1 ms per
  request plus bytes at 100 MiB/s — not observed service. No range-source
  selector exists for `validate_source`, and adding one before this change would
  have cost 82.7 s per sample on the flagship fixture. That asymmetry is the
  reason the counts carry the argument and the model only prices it.
* **No cold-cache, physical-I/O, network or device result.** The `file` legs run
  over a warm page cache; what they measure is `pread`/`fstat` call count, not
  physical I/O.
* **The whole-sheet walk and the text projection are unchanged, not improved.**
  0630's item 18 expected a range-source improvement there and this change does
  not deliver one; see the next paragraph for why and what would.
* **The 64 KiB ceiling and the 512-byte first step are 0568's numbers, adopted,
  not re-derived here.** No sweep of alternative ceilings was run.
* **One caller.** `litchi-doc` and `litchi-ppt` build no stream cursors, so the
  wrapper has exactly one user today.

**Left open, and named for the queue.** The 16,061 small requests 0627 measured
on the `54016.xls` whole-sheet walk are **per-shared-string cursor
constructions**, each reading one string's bytes at a random offset in the SST
region, 4,444 of them jumping backwards. Attacking them needs a bounded,
locality-aware view of the shared-string region that survives across
resolutions — a different mechanism from this one, in `litchi-xls`, and one that
must answer what it costs a workbook whose SST does not fit its budget. It is
not proposed here.

**Rejected on the way.** Putting the window inside `SharedOleStreamCursor` is
implemented, tested and measured, and rejected: it costs the whole-sheet walk
+1.41% cycles and +6.20% p50 because the shared-string resolver builds 16,105
cursors on that walk. The wrapper is the same mechanism with no cost to callers
that do not use it.

## Retained evidence

[`results/change-0636/README.md`](results/change-0636/README.md) — the
attribution probe and its manifest template, the two corpus sweeps and their
diff, the `perf stat` and callgrind pairs with their runners, every timing leg
verbatim with two independent repeats of the two scenarios that moved, the two
range-source reports and their comparison, the gate tails, the decision record
and the log paragraphs.
