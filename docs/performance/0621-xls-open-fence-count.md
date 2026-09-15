# 0621: the XLS text projection observes the source once per read, not four times per shared string

Status: retained. `performance_claim: none` — this record carries counted
invariants, syscall counts, instruction counts and paired medians as evidence,
and registers no claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This record takes change [0587](0587-remaining-opportunity-survey.md) item
**CORE-2** (rank 34, "collapse the 25 per-open `version()` fences") together with
the fence half of item **XLS-4** (rank 33, "per-string and per-row freshness
fences"). It answers both, and they do not come out the same way: CORE-2 is
**falsified on its own terms** and XLS-4's fence half is confirmed, with the
survey's model of it accurate to within 0.04%.

## What was removed

Every site below was analysed in the style of change
[0563](0563-opc-single-warm-part-observation.md): for each observation, what
mutation is detected between it and its neighbours, and whether that detection
is a strict subset of another observation's. The complete per-site table is in
[the analysis section](#the-per-site-analysis); the measured attribution behind
it, one bucketed backtrace per observation, is
[`sites/`](results/change-0621/sites) in the packet.

Ten sites stop observing on the path that succeeds. **Five of the ten are
relocations rather than deletions**: the observation survives on the error branch
where it outranked another failure, or inside the probe whose read made it
load-bearing.

**`crates/litchi-xls/src/workbook/source.rs`**

- **`resolve_shared_string`, three sites, 48,187 of the 89,789 observations a
  `54016.xls` text projection took.** The resolver observed the source before
  searching its segment table, again before each chunk read, and again after
  decoding. It consumes source bytes only through
  `SharedOleStreamCursor::read_exact`, which change
  [0558](0558-ole2-single-read-fence.md) already leaves observing once after
  every read; the resolver's own three added nothing on the path that succeeds —
  the first sits between the scan's last read fence and the first chunk read, the
  second is the leading fence of a read whose trailing fence is the cursor's, and
  the third sits behind an in-memory decode. All three are removed. They were
  load-bearing for **error precedence**, so the function is now wrapped the way
  `SharedOleFile::finish_stream_range` is: one observation on the failure path,
  which promotes `SourceChanged` over the locator, allocation, chain and decode
  errors a mutation can provoke. Change
  [0317](changes/0317-opc-source-read-error-precedence.md) fixes that ordering
  and it is preserved on every failure branch, including two
  (`SST entry locator is outside its segments` and the `chunks` allocation
  failure) that no fence covered as precisely before.
- **`write_text_sheet`'s per-row fence, 5,096 observations on the same
  projection.** The row loop emits rows out of the map `scan_text_sheet` has
  already collected and fenced; it reads no source byte. Every byte it emits goes
  through `SourceCheckedTextSink`, which observes immediately **before** the
  write and again immediately after, so the row fence sat between two in-memory
  steps and the sink re-proved it a moment later, before any byte left. The
  source half is removed. The **cancellation** half is kept and split into
  `check_text_cancellation`, because cancellation is not in the same position: a
  row whose object `SequentialTextWriter::write_object` skips reaches no sink
  call at all, so the row loop must keep its own cancellation granularity.
  `write_text_sheet` no longer needs the owner and no longer takes it.
- **The duplicate capture in `SourceBackedWorkbook::from_shared_ole_file_with_limits`.**
  `cfb.source_version()` observes the retained source and refuses when it has
  moved; the `ensure_current_parts` call on the next line observed the same
  object, against the same value, with nothing between. Change
  [0560](0560-xls-single-observation-freshness.md) collapsed exactly this pair
  inside `ensure_current_parts` itself. The second call is removed rather than
  the first, so the error identity on both branches is the one the code already
  produced: `SourceBackedError::SourceChanged` with the same payload, and
  `SourceBackedError::Cfb(OleError::Io(..))` when the observation itself fails.
- **`select_workbook_stream`'s leading fence.** `SharedOleFile::stream_len`
  resolves a name against the directory tree captured and validated at open and
  consumes no source byte, so the fence at the head proved what each of the four
  exits proves. Each exit keeps its own, which is where it matters: a changed
  source still outranks `WorkbookStreamMissing` and the mapped CFB error. This is
  the shape change 0560 gave the retained-metadata helpers — fence once, after
  the value is produced.

**`crates/litchi-cfb/src/shared.rs`**

- **The middle observation of `open_source_with_limits`.** The open captured the
  version, observed the length, observed again, parsed, and observed a third
  time. The length is bracketed by the capture and the post-parse comparison
  exactly as the parse is, so the middle observation is a strict subset of the
  third on the path that succeeds. It is **relocated**, not deleted, onto the two
  branches that leave before the post-parse comparison can run — a failed length
  observation, and a length over `max_input_bytes` — where without it a
  length-derived failure would be reported while the source was already stale.
  `Self::refuse_if_changed` is the open-path form of `check_source_version`,
  which cannot be used because no `SharedOleFile` exists yet to hold the captured
  version.

**`crates/litchi/src/detection_smart/detected.rs`**

- **The leading fence of the signature read** in
  `detect_workbook_source_path_with_limits`. The fence after that read is also
  the fence the length is compared under, because every branch that rejects
  `source_length` runs after it. The leading one is removed from the path that
  succeeds and relocated onto the read's failure path, so a truncation racing the
  read still surfaces as `SourceChanged`.
- **The two unconditional probe fences.** The ODS mime probe and the ODS catalog
  probe each ran under a condition (`zip_magic`, then `is_ods`) and were followed
  by an unconditional fence. Each fence now sits **inside** its probe. When a
  probe reads, its bracket is exactly what it was; when it does not — which is
  every OLE2 workbook — the fence was re-proving the signature read's fence. This
  also removes the fence from the three `#[cfg]` arms that compute `ordinary_ods`
  without probing anything.

## What was kept, and why

The count cannot fall much further without giving up something ADR 0005 needs,
and the kept sites are the interesting half of the analysis.

- **One observation after every source read** (16 per `54016.xls` open from
  `GlobalsBuffer::ensure`, 16,093 per text projection from
  `SharedOleStreamCursor::read_exact`). This is the fence that makes a
  mid-operation truncation surface as `SourceChanged` rather than as a garbled
  read, and change 0558 already reduced it from two to one. **It is the floor of
  the open count**, and the reason CORE-2's target is unreachable — see below.
- **Both observations in `SourceCheckedTextSink::write`**, 20,382 on a
  `54016.xls` projection and the largest single site left. The trailing one
  detects a mutation spanning the caller's write. The leading one is what
  *stops* the stream: the trailing check records its failure but returns `Ok`
  from `write`, so without a leading check on the next call the writer would emit
  the whole remaining document and refuse only at the end. Neither is a subset of
  the other, and the test
  `a_mutation_from_inside_the_output_sink_stops_the_text_stream` pins it.
- **The open's trailing fence** (`inner.ensure_current()` at the end of
  `from_shared_ole_file_with_limits`). Strictly it re-proves the last globals
  fill's trailing fence, but it is the public operation boundary, and keeping it
  makes "the open is bracketed" true independently of whether `parse_globals`
  read anything. The same argument does not apply to `resolve_shared_string`,
  which is an internal step of a scan whose boundary is `finish_scan` or
  `write_text_to_impl`'s tail.
- **`query_cell`'s leading fence and its out-of-range twin** (two observations
  with nothing between them, on a branch that reads nothing). Analysed as
  removable and deliberately not changed: it is outside this record's scope, the
  branch is unreachable from any harness selector, and two other changes are in
  flight in `litchi-xls`.

## The per-site analysis

Every observation reachable on the three paths, in the order the reader takes
them. "Detects" is the mutation that observation and no earlier one can see;
"subset of" names the later observation that sees the same mutation with nothing
consumed from the source in between. Counts are the measured attribution for one
`54016.xls` full text projection, from
[`sites/54016-full-text-before.txt`](results/change-0621/sites/54016-full-text-before.txt).

### `crates/litchi-cfb/src/shared.rs`

| site | count | detects | verdict |
| --- | ---: | --- | --- |
| `open_source_with_limits`, capture | 1 | opens the bracket for the length and the parse | **kept** — it is the capture |
| `open_source_with_limits`, after `len()` | 1 | a mutation spanning the length observation; outranks the `len()` failure and `LimitExceeded` | **relocated** onto those two branches: a subset of the post-parse comparison wherever that comparison is reached |
| `open_source_with_limits`, after parse | 1 | a mutation spanning the parse | **kept** — closes the bracket |
| `SharedOleStreamCursor::read_exact`, trailing | 16,093 | a mutation spanning the read whose bytes it publishes | **kept** — this is the fence that makes a truncation a refusal rather than a garbled read |
| `read_stream_range_hinted`, trailing (`GlobalsBuffer::ensure`) | 16 | the same, for each globals fill | **kept**, same reason |

### `crates/litchi-xls/src/workbook/source.rs`

| site | count | detects | verdict |
| --- | ---: | --- | --- |
| `from_shared_ole_file_with_limits`, `cfb.source_version()` | 1 | the snapshot's opening observation and its comparison | **kept** |
| `from_shared_ole_file_with_limits`, `ensure_current_parts` | 1 | nothing: same object, same expected value, nothing consumed between | **removed** — strict subset of the line above |
| `select_workbook_stream`, head | 1 | nothing: `stream_len` resolves the captured directory tree | **removed** — strict subset of each of the four exits |
| `select_workbook_stream`, four exits | 1 | a mutation spanning the lookups; outranks `WorkbookStreamMissing` and the mapped CFB error | **kept** |
| `from_shared_ole_file_with_limits`, trailing | 1 | re-proves the last globals fill's fence | **kept** — the public operation boundary |
| `resolve_shared_string`, before the segment search | 16,055 | nothing on the success path; outranks the locator, cursor and allocation errors | **removed**, precedence moved to the failure wrapper |
| `resolve_shared_string`, per chunk | 16,077 | nothing: the leading fence of a read whose trailing fence is the cursor's | **removed** (change 0558's argument, one level up) |
| `resolve_shared_string`, after a successful decode | 16,055 | nothing: only an in-memory decode since the last chunk's fence | **removed** |
| `resolve_shared_string`, after a failed decode | 0 on this fixture | outranks the decode error | **replaced** by the wrapper's single failure-path observation, which also covers the three above |
| `write_text_sheet`, per row | 5,096 | nothing: the row is built from the collected map, and the row's first sink write observes a moment later, before any byte leaves | **removed** (source half); the cancellation half is kept because a skipped object reaches no sink call |
| `SourceCheckedTextSink::write`, leading | 10,191 | stops the stream: the trailing observation only records its failure and returns `Ok` | **kept** |
| `SourceCheckedTextSink::write`, trailing | 10,191 | a mutation spanning the caller's write | **kept** |
| `text_impl` and `write_text_to_impl` brackets, per sheet and per operation | 7 | the operation boundaries, and the tail that promotes a source failure over any conversion error | **kept** |
| `query_cell`, leading and out-of-range | 2 per query | the second re-proves the first with nothing between | **analysed, not changed** — out of this record's scope |
| `scan_worksheet`'s `finish_scan`, `visit_worksheet_cells`, `metadata` | 1 each | operation boundaries | **kept** |

### `crates/litchi/src/detection_smart/detected.rs`

| site | detects | verdict |
| --- | --- | --- |
| capture | opens the bracket | **kept** |
| after `len()`, before the signature read | the leading fence of that read; the length is compared under the fence after it, because every branch that rejects the length runs later | **removed** from the success path, **relocated** onto the read's failure path |
| after the signature read | a mutation spanning that read, and the length | **kept** |
| after the ODS mime probe | load-bearing only when `zip_magic` made the probe read | **moved inside the probe** |
| after the ODS catalog probe | load-bearing only when `is_ods` made the probe read | **moved inside the probe** |
| `classify_ole_host_stream_from_shared` | a mutation spanning the four directory lookups | **kept** |

## Why it is sound

**ADR 0005.** The rule is that a mutation during a read returns `SourceChanged`.
Every path that consumes a source byte is still bracketed by two successful
observations of the captured version — the most recent earlier one taken by the
reader, and the read's own trailing fence — which is the discipline change 0558
established and documented on `read_exact` and `read_stream_range`. Nothing this
change removed sat between a read and the bytes it published. ADR 0005 does not
require a `statx` per stream read, and after this change the XLS text path takes
none beyond the reader's own.

**Error identity and precedence.** Every removed observation that outranked
another error keeps that precedence, by relocation rather than by deletion:
the CFB open's two early returns, the facade's read-failure path, and
`resolve_shared_string`'s new failure wrapper. `write_text_to_impl` already
promoted a source failure over any conversion error at its tail
(`take_source_text_failure(..).or_else(..).or_else(|| self.inner.ensure_current().err())`
runs before `conversion?`), and that is what makes the text path's intermediate
precedence fences redundant rather than merely cheap.

**One stated consequence.** In `open_source_with_limits` the caller-limit
construction (`OleFileLimits::new(limits.max_input_bytes())?`) now runs before
the first comparison that can see a mutation, so a caller who passes an invalid
ceiling *and* races a mutation receives the limit error where the mutation was
reported before. The limit error is source-independent and deterministic — that
caller receives it with no mutation at all — so no refusal is traded for a
partial result.

**Reverted-mutation sampling.** `FileVersionPolicy` already documents that a
transition fully reverted between observations is not visible, and change 0558
accepted that halving the observations halves the sampling of that latch. This
change reduces it further on the same terms and for the same reason: a mutation
that persists is observed by the next fence, and by every later one.

**Contracts untouched.** No `unsafe`, no limit changed, no validation moved,
relaxed or deferred, no I/O added or removed, no allocation added on any hot
path, no public type changed. The number and size of physical reads is identical
in every measured cell. Cancellation granularity is unchanged: every removed
fence that carried a cancellation check keeps it.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Base `1e41983213dc378c13774ed7038c51faf231977f`, branch
`perf/0621-xls-open-fence-count`. Both legs built `--release
--locked`; the before leg from the shared read-only checkout of the base with its
own `CARGO_TARGET_DIR`. CPU 16 pinned; seven other agents active on other cores;
host-wide quiescence is not established. Full provenance and binary digests in
[the packet README](results/change-0621/README.md).

### Counted observations (primary evidence)

`tools/perf-baseline`'s `xls_source_attribution`, `--mode file-source`, one
sample. `read_calls`, `read_bytes` and the projection digest are **identical in
every cell**; the `owned-readat` control reports the same observation counts, so
these are logical observations and not an artifact of the file adapter.

| fixture | operation | before | after | change |
| --- | --- | ---: | ---: | ---: |
| `54016.xls` | open | 25 | **22** | −12.0% |
| `54016.xls` | list worksheets | 25 | **22** | −12.0% |
| `54016.xls` | one cell | 43 | **40** | −7.0% |
| `54016.xls` | whole-sheet walk | 64,307 | **16,117** | **−74.9%** |
| `54016.xls` | full text | 89,789 | **36,503** | **−59.3%** |
| `ConditionalFormattingSamples.xls` | open | 29 | **26** | −10.3% |
| `ConditionalFormattingSamples.xls` | one cell | 38 | **35** | −7.9% |
| `ConditionalFormattingSamples.xls` | full text (typed refusal) | 1,936 | **916** | −52.7% |
| `SimpleMultiCell.xls` | open | 16 | **13** | −18.8% |
| `SimpleMultiCell.xls` | full text | 59 | **49** | −16.9% |

Where the 89,789 went, from the per-site attribution
([`sites/54016-full-text-before.txt`](results/change-0621/sites/54016-full-text-before.txt)):

| site | before | after |
| --- | ---: | ---: |
| `SourceCheckedTextSink::write`, two per written object | 20,382 | 20,382 |
| `SharedOleStreamCursor::read_exact`, one per read | 16,093 | 16,093 |
| `resolve_shared_string`, three per shared string | 48,187 | **0** |
| `write_text_sheet`, one per row | 5,096 | **0** |
| `GlobalsBuffer::ensure`, one per fill | 16 | 16 |
| operation and open brackets | 15 | 12 |

The survey modelled the resolver and row fences at "**≥48,165** `fstat` plus one
per row"; measured, they are **53,283** of 89,789, and the whole projection took
**89,794** `statx` where the survey had no figure at all.

### Syscalls

`strace -f -c`, whole child, one and eleven samples differenced and divided by
ten, so these are per-operation.

| fixture | mode | operation | `statx` before | after | `pread64` |
| --- | --- | --- | ---: | ---: | --- |
| `54016.xls` | file-source | open | 30 | **27** | 40 → 40 |
| `54016.xls` | facade-file | open | 36 | **31** | 42 → 42 |
| `54016.xls` | file-source | full text | 89,794 | **36,508** | 16,145 → 16,145 |
| `ConditionalFormattingSamples.xls` | file-source | open | 34 | **31** | 53 → 53 |
| `ConditionalFormattingSamples.xls` | facade-file | open | 40 | **35** | 55 → 55 |

`pread64` is identical in every cell.

### Instructions

Callgrind isolation pairs, one and six samples after one warm-up, differenced
and divided by five.

| case | before | after | change |
| --- | ---: | ---: | ---: |
| file-source open | 5,099,495 | 5,110,684 | +0.22% |
| facade-file open | 5,084,091 | 5,099,531 | +0.30% |
| file-source full text | 426,888,900 | **400,986,804** | **−6.07%** |

The two opens are flat, and slightly adverse: valgrind does not simulate the
syscall, so a removed `statx` is nearly free in instructions and what remains is
code placement. The text projection's −25.9 M instructions are the mutex, the
metadata fingerprint, the comparison and the call frames of 53,286 removed
observations.

### Paired timing

Fresh child per leg, ASLR disabled, CPU 16, 3 warm-ups, 60 samples (40 for the
text projection), order A1 B1 B2 A2 with two further before legs in the same
window for the A/A floor. Both directions reported.

| case | statistic | before | after | direction 1 | direction 2 | A/A floor |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| full text, file-source | p50 | 48.73 ms | **35.31 ms** | **−27.55%** | **−28.32%** | −3.71% |
| full text, file-source | mean | 49.13 ms | 35.25 ms | −28.24% | −27.92% | −3.19% |
| full text, file-source | p95 | 53.28 ms | 37.05 ms | −30.45% | −24.96% | −4.57% |
| full text, file-source | p99 | 53.43 ms | 37.25 ms | −30.28% | −26.82% | −4.60% |
| open, file-source | p50 | 212.2 µs | 210.6 µs | −0.76% | −0.70% | +0.31% |
| open, facade-file | p50 | 214.2 µs | 210.7 µs | −1.60% | −0.19% | +0.12% |
| one cell, file-source | p50 | 861.6 µs | 850.2 µs | −1.33% | −1.41% | +0.10% |

The text projection is the only case that clears the floor, and it clears it by
six to eight times. The accounting closes: 53,286 removed observations at the
167 ns change 0587 measured for `FileSource::version()` on this host is 8.9 ms,
the removed instructions are about 4 ms of the remainder, and the observed move
is 13.4 ms. **No case is adverse in both directions by more than 5%.** One
comparison is adverse in both directions at all: `facade-file` open p99, +3.17%
and +0.43% against a +1.91% floor, which is reported rather than excluded.

The three sub-millisecond cases move −0.19% to −1.60% at p50, consistently
negative and outside their own in-window A/A floors (+0.10% to +0.31%, which are
this tight because every leg is a fresh child with ASLR disabled) but well inside
the program's stated host floor of about 4% at p50. They are reported as
evidence and **no latency claim is made for them**. That is the expected result
and the reason CORE-2 was ranked as a count item: the survey modelled 25-36
observations at 167 ns as 1-1.5% of an open, and the measured move is of that
order.

### CORE-2 is falsified; XLS-4 is confirmed

0587 set CORE-2's falsification test as "moving fences changes any
`SourceChanged` outcome in the change-under-read tests, **or the count cannot
drop below about 10 per open**". The first does not happen. The second does: the
open is now 22 observations on `54016.xls`, of which **16 are the trailing fence
of the 16 globals fills**, 2 are the CFB open's own bracket, and the remaining 4
are the snapshot's capture, `select_workbook_stream`'s exit, the open's trailing
fence and the `worksheet_count` query the harness's `open` operation makes.
**18 are structural**: nothing below that is reachable by moving a fence, because
it would mean removing the per-read fence ADR 0005 needs and change 0558 already
halved. The remaining lever on that number is to issue fewer reads — 16 globals
fills is what item XLS-6 and its design record are about — not to move fences.
CORE-2 as ranked, "collapse the 25 to one per operation boundary", **cannot be
done**, and this record says so rather than reporting the 12% it did get as if it
were the item.

0587 set XLS-4's falsification test as "`strace -c` on a `from_path` extraction
puts fences under 5% of wall time". It does not: the removed fences alone are
18% of the before extraction's wall time by the 167 ns rate, and the measured
p50 move is 27.6%. XLS-4's fence half is confirmed, and the survey's model of it
was **accurate**: it predicted "≥48,165 `fstat` (3×16,055) plus one per row",
which is 53,261 for this fixture's 5,096 rows against a measured 53,283, within
0.04%. What the survey did not have is the denominator — the whole projection
takes **89,794** `statx`, so the fences it enumerated are 59% of the operation's
observations and the rest belong to the cursor and the output sink. Its other two
parts — the declared-rectangle walk and the three hash operations per
`SourceTextSheet::insert` — are untouched here.

## Correctness evidence

**Corpus differential.** Every `.xls` fixture under `test-data/ole/xls` and
`test-data/poi/test-data/spreadsheet` — 109 files — was put through both the full
text projection and the whole-sheet walk on both legs, 218 cells. **All 218 agree
exactly**: the same SHA-256 of the projected text or of the walked cell
sequence, or character-identical typed refusals for the 19 cells the reader
refuses, and the same `read_calls` and `read_bytes`. Observations over the
corpus fall 230,619 → 91,611 (−60.28%). Outputs and the analysis script are in
[`corpus/`](results/change-0621/corpus).

**Tests added.** Nine: six in `crates/litchi-xls/tests/source_backed.rs` and
three in `crates/litchi-cfb/src/shared.rs`. Three are counted invariants that
fail if a fence is put back; six are preservation tests that pass on both legs:

- `a_text_extraction_observes_the_source_once_per_shared_string_read_and_no_more`
  measures the **slope**, not the intercept: eight further shared-string cells on
  one row cost eight reads and eight observations. With the resolver's fences
  restored the same measurement is eight reads and **32** observations.
- `a_text_extraction_observes_the_source_once_per_read_and_twice_per_written_object`
  pins the composition of an open and a projection on `Simple.xls` at
  (12, 4, 21); with the fences restored it is (15, 4, 27).
- `an_open_observes_the_source_twice_around_its_parse` pins the CFB open at two;
  with the fence restored it is three.
- `a_mutation_in_any_observation_window_of_a_text_extraction_is_refused` and
  `..._of_an_open_is_refused` are the tests this record's brief asks for. Every
  fence that was removed sat between two observations the reader still takes, so
  the sweep places a mutation in **every one of those windows in turn** — the
  `CountingSource` gains a `bump_after_observation(n)` hook — and requires the
  same typed refusal from all of them. The control at the end mutates after the
  operation's last observation and requires success, so the sweep cannot pass by
  refusing everything.
- `a_mutation_from_inside_the_output_sink_stops_the_text_stream` mutates the
  source from inside the caller's `Write` and requires that the stream stop at
  that write rather than emit the rest of the document.
- `a_changed_source_outranks_a_missing_workbook_stream`,
  `an_open_refuses_a_changed_source_over_its_input_limit` and
  `an_open_refuses_a_changed_source_over_a_failed_length` pin the three relocated
  precedences, each with a stable-source control that requires the original
  error.

**Gates.** `cargo fmt --all --check` clean. Clippy clean on `litchi-cfb` and
`litchi-xls` with all targets. `cargo doc --no-deps` clean on `litchi-cfb`,
`litchi-xls` and `litchi`. `cargo test -p litchi-cfb` 356 passed, 0 failed;
`cargo test -p litchi-xls` 1,396 passed, 0 failed across 72 test binaries.
`litchi` carries one pre-existing test failure and four pre-existing warnings
under the feature set that compiles the facade's XLS route, all reproduced
verbatim on the untouched before checkout; see
[`gates.txt`](results/change-0621/gates.txt).

## Validation preserved

Nothing was added to or removed from any validation. The CFB directory tree,
allocation chains and final-sector rules are still parsed and validated at open
under the same limits and compared against the same captured version
afterwards; `validate_sheet_offsets`, `validate_worksheet_bof`,
`Formatting::validate_cell_xf` and the SST locator and segment checks are
untouched; the text output's byte and object limits are untouched. The 19 corpus
cells that the reader refuses are refused with character-identical messages on
both legs.

## Limitations

No cold-cache, physical-device, remote or range-source, peak-RSS, allocation,
concurrency-scaling or cross-platform result is claimed. The latency figures are
from one host with seven other agents active and no proven quiescence; the
counted observations, syscalls and instruction counts are not.

The count claim's relevance on a network filesystem, where a `statx` is a round
trip rather than 167 ns, is **modelled**, not measured: this host has no real
range source and change 0587 recorded the same blocker.

The `facade-file` mode measures the facade's source-backed XLS route only because
`tools/perf-baseline` enables `litchi/xls` through its `xls-source-attribution`
feature; the whole-child `statx` figures for it are not decomposed per site,
because the release binary carries no symbols for `strace -k` to resolve. The
facade's PPT route (`detect_presentation_source_path`, eight fences) was read and
is not changed: it is behind a different `#[cfg]`, has no selector, and is not on
any path this record measures.

Three sites analysed as removable are deliberately left: `query_cell`'s leading
fence and its out-of-range twin, and the open's trailing fence. The reasons are
in [What was kept](#what-was-kept-and-why). The rectangle walk and the three hash
operations per `SourceTextSheet::insert` — the other two parts of 0587's XLS-4 —
are untouched, and `SequentialTextWriter` still issues two `write` calls per row
(a separator and a value), which is what makes the sink's 20,382 observations the
largest site left on the text path; folding those two writes into one is the
named follow-on and is not a fence question.

## Retained evidence

[`results/change-0621/README.md`](results/change-0621/README.md) — per-site
attribution for both legs, counts, `strace` and callgrind outputs, the 24 timing
children, the corpus differential, every script, and the source of the
backtrace-bucketing probe that produced the attribution.
