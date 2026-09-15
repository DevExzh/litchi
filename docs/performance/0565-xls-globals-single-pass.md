# 0565: one pass over the XLS globals, each byte read once

Status: retained. `performance_claim: none` — this record carries deterministic
counters, isolated syscall counts and paired latency, not a registry claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## The decision this record makes

Change [0564](0564-xls-open-read-attribution.md) attributed 621 of the 655
positional reads of a source-backed XLS open to `parse_globals`'s header-only
pre-pass — one four-byte read per BIFF globals record, followed by a bulk read
of `[0, global_end)` that fetched every one of those header bytes again — and
recorded that removing it conflicts with four contracts that tests assert
deliberately. It enumerated what a fused pass would have to establish and, per
`docs/GOAL.md`, stopped rather than implementing through the conflict.

This record establishes them and implements the pass. The decision was taken
explicitly on three grounds. The four contracts are test-level statements of the
previous implementation's read shape, not ADR requirements, and no accepted ADR
is weakened. `docs/GOAL.md` workstream A asks for exactly this: "bounded
read-ahead and range coalescing based on measured access patterns", and avoiding
"request amplification from many tiny remote reads". And every invariant those
contracts protect survives in a form that is still enforced by a test.

## What the scan does now

One pass fills a retained buffer holding stream bytes `[0, filled)`, and the
buffer, truncated to `global_end`, is handed to the unchanged semantic pass. The
bulk re-read is gone.

**An exact prologue covers the first four records.** Each costs one header read
plus one read carrying that record's payload *and the next record's four-byte
header*. `[MS-XLS]` 2.1.7.20.1 defines the globals substream as
`GLOBALS = BOF [WriteProtect] [FilePass] [Template] ...`, so a `FilePass` record
occupies one of the first few positions. Every encrypted fixture in the
repository carries it at record index 1, byte offset 20. Because the coupled
read delivers the next header, the safe prefix is in fact **five** records: a
`FilePass` at index 4 is also refused before any fill is issued.

**From the fifth record on the scan fills in windows**, starting at 512 bytes —
one CFB sector — and doubling to a 64 KiB cap. Every fill is clamped by the
stream length, by `max_global_bytes`, and by the smallest `BoundSheet8` stream
position framed so far.

## Why a short prologue and not an exact phase

The first draft windowed only after the contiguous `BoundSheet8` run had ended,
so the clamp was always known before any fill and nothing past the globals end
could ever be read. A survey of 104 container fixtures killed that design: the
first `BoundSheet8` record is late in record order but early in byte order — for
`ConditionalFormattingSamples.xls` it is record 526 of 621 but byte 20,783 of
551,377.

Modelled over the 54 fixtures that carry a workbook stream, the three candidate
schedules cost:

| Schedule | Logical reads | Over-read past the globals end |
| --- | ---: | ---: |
| today | 8,543 | 0 |
| exact until the clamp is known | 7,872 | **0** |
| prologue then windows (implemented) | **434** | 32,532 B, worst 3,949 |

Eliminating the over-read costs essentially the entire benefit. The survey and
both alternatives are reproducible from one generator in
[`corpus-survey/`](results/change-0565/corpus-survey/README.md).

## Accepted contract changes

| Contract in change 0564 | What replaces it |
| --- | --- |
| Open's range list is exactly one 4-byte range per globals header, then one bulk range | Every byte of `[0, global_end)` is read **exactly once**, no two reads overlap, and the read count is bounded by the prologue plus the fill count |
| A `FILEPASS` payload is never read | Exact for the first five records, which is where every encrypted fixture in the repository carries it. A later `FILEPASS` may have payload bytes resident in a fill buffer; they are never framed, interpreted, logged, put in an error message or published |
| A skipped payload is never read | Unchanged. That contract belongs to the worksheet scan, which this change does not touch |
| Open never touches a sheet body | Every fill is clamped at the smallest `BoundSheet8` position **once one has been framed**. Before that a fill may reach past the globals end; the read is bounded by one fill and by the stream, and the bytes are truncated away |

Three further changes are accepted and recorded rather than discovered later:

- **`max_global_bytes` remains a plus-four bound, not a hard bound**, exactly as
  before: the four header bytes that prove a record crosses the limit are read,
  then the limit error is reported with the record's real end as `observed`. A
  new test pins it.
- **Read coupling changes error precedence in two places.** In the fill phase an
  I/O or `SourceChanged` failure at a later offset can be reported before a
  `FILEPASS` or limit error of an earlier record in the same fill. In the
  prologue the same applies within one record pair. The per-record check order
  is unchanged everywhere.
- **The clamp is a running minimum**, not the minimum over the whole stream.
  `[MS-XLS]` 2.4.28 imposes no ordering on `lbPlyPos`; the survey finds the
  positions ascending in 54 of 54 fixtures, which is empirical rather than
  normative. A running minimum is still a correct bound because it only lowers.

## A third site held the old contract

The perf harness independently encoded "an open reads zero worksheet bytes" as a
read-locality gate, on a corpus purpose-built to test locality. The change trips
it: that corpus has only 1,483 bytes of globals, so the last fill is issued
before any `BoundSheet8` exists and reads **93 bytes** of an unselected
worksheet body.

The gate is bounded rather than removed. It now requires that an open read **no
byte of the selected worksheet at all** and at most one window of any worksheet
body, which still fails an open that materializes a worksheet — the regression
it was written to catch, since the corpus worksheets are far larger than one
window. Both measurement binaries were rebuilt from trees carrying that identical
gate, so the two legs differ only in the library change.

**Correction (change [0619](0619-harness-xls-lifecycle-assertion.md)).** A
fourth site held the old contract and was not updated. The gate
`validate_xls_source_locality` was bounded here, but the two published
evidence booleans it sits beside — `open_reads_zero_worksheet_payload` and
`selected_query_reads_only_selected_worksheet` — are still computed with the
strict `== 0`, and `tests::xls_source_backed_lifecycle_selectors_are_matched_and_local`
asserted the pre-change contract at three sites through them. That test went
red at this commit (`c1d2caf85`; it passes at the parent `6b13261e5`) and
stayed red for 54 records, its panic poisoning the shared allocation-metrics
mutex so six further harness tests failed as cascades. Change 0601 reported
the failure as pre-existing and pointed at change 0595's area; change 0619
bisected it back to here and measured the cause to be exactly the 93-byte
over-read this record already documents — no library read moved. This record's
measured effect, its counters and its correctness evidence are unaffected. The
gap was structural: `tools/perf-baseline` is a separate Cargo project, so the
gate list used here — 1,351 `litchi-xls` tests and one facade test — could not
reach the harness's own suite.

## Measured effect

Corpus `test-data/ole/xls/ConditionalFormattingSamples.xls`, 1,402,368 bytes.
Environment and identities are frozen in
[`plan.json`](results/change-0565/plan.json); CPU 17 pinned, ASLR disabled for
the selector matrix. Host quiescence is not established.

### Deterministic counters

Every counter was identical across all 100 samples of every attribution child.

| Operation | logical reads | read bytes | `version()` calls |
| --- | ---: | ---: | ---: |
| open | 655 → **53** | 567,685 → 565,201 | 631 → **29** |
| list | 655 → **53** | 567,685 → 565,201 | 631 → **29** |
| one cell | 921 → **319** | 569,398 → 566,914 | 902 → **300** |

The globals scan itself falls from 630 reads to 28. Bytes fall by 2,484, which
is exactly the header bytes the old bulk read fetched a second time.

### Syscalls

Isolated by differencing `strace -f -c` runs at 1 and 11 samples with one warmup
and dividing by ten, as change 0564 did.

| Operation | `pread64` | `statx` |
| --- | ---: | ---: |
| open | 655 → **53** (−91.9%) | 636 → **34** (−94.7%) |
| list | 655 → **53** (−91.9%) | 636 → **34** (−94.7%) |
| one cell | 921 → **319** (−65.4%) | 907 → **305** (−66.4%) |

The four-byte read share collapses. In a two-open capture, four-byte reads fall
from 1,242 to **4**, and the size histogram becomes the doubling schedule:

```
before  {4: 1242, 512: 48, 840: 2, 1536: 2, 43473: 2, 49152: 2, 65536: 14}
after   {4: 4, 6: 4, 20: 2, 512: 50, 840: 2, 1024: 2, 1536: 2, 2048: 2,
         4096: 2, 8192: 2, 15912: 16, 16384: 2, 16856: 2, 27561: 2, 49624: 14}
```

### Paired latency

An A1/B1/B2/A2 matrix over ten XLS selectors with ASLR disabled, 5 warmups and
60 samples per child: 40 children, 14 case and corpus rows, 56 statistic
comparisons, **27 improving in both directions**.

| Selector | p50, first direction | p50, second direction |
| --- | ---: | ---: |
| `xls_source_backed_open_list_worksheets` | −11.66% | −3.82% |
| `xls_source_backed_open_one_cell` | −8.87% | −6.27% |
| `xls_source_backed_open` | −7.48% | −2.29% |
| `xls_owned_source_open_one_cell` | −7.34% | −4.03% |
| `xls_owned_source_open_list_worksheets` | −3.67% | −9.41% |
| `xls_owned_source_open` | −2.38% | −5.91% |
| `xls_semantic_full_cell_scan` (`xls-large`) | −2.48% | −5.00% |

**The owned-source selectors were predicted to move within noise. They do not,
and that prediction is recorded as falsified.** The before-capture had already
shown why: an owned source performs the same 655 logical reads and 567,685
logical bytes with **zero** syscalls. So these rows isolate per-logical-read cost
from per-syscall cost, and they show the change removes both. The attribution
binary's own timing agrees: its median open falls 72.2% on a file source and
47.8% on an owned source.

### Review triggers

Two comparisons are adverse in both directions by more than 5%, both on the same
selector at nanosecond scale:

| Case | Statistic | Change | Absolute |
| --- | --- | ---: | --- |
| `xls_semantic_list_worksheets` `xls-large` | p50 | +16.67% / +16.67% | **60 ns → 70 ns** |
| `xls_semantic_list_worksheets` `xls-large` | mean | +9.58% / +22.83% | 87 ns → 95 ns |

A 60 ns median moves by one 10 ns clock tick for 16.67%, so these cells cannot
distinguish a real regression from quantization. They are reported rather than
excluded. Change 0560 recorded the same artifact on the same harness. An A/A dry
run on this host at 20 samples drifted +1.7% to +3.2% at p50 and +11.0% at p99
with one binary against itself, so tail movements below roughly 10% carry no
information here.

## Correctness evidence

`litchi-xls` passes **1,351 tests with zero failures**. Tests added, each
confirmed to fail against the pre-change implementation:

- `source_backed_open_reads_each_global_byte_once` asserts every globals byte is
  read exactly once, over-read exactly zero, and no read at or beyond the
  smallest sheet position, on a fixture large enough that the window bound
  binds rather than the stream.
- `mini_stream_workbook_globals_are_read_once_each` covers a workbook stream
  resident in the CFB mini stream, so the fills go through the mini-stream range
  reader.
- `corrupt_bound_sheet_positions_terminate_the_globals_scan` pins that a corrupt
  `lbPlyPos` behaves exactly as before the change and that the scan still stops
  short of the stream end after the clamp is dropped.
- `every_bound_sheet_lowers_the_globals_fill_clamp` builds a synthetic fixture
  whose sheet positions descend and where a fill boundary genuinely lands
  between two `BoundSheet8` records. It is mutation-checked: keeping only the
  first position makes exactly this test fail.
- `mid_globals_byte_limit_reads_at_most_one_header_past_the_limit` reaches the
  limit through the fill phase. It fails pre-change **only** on the byte count;
  the typed error, its `observed` value and the plus-four containment all pass
  pre-change, which is the confirmation that the error identity is unchanged.

Two tests pin preserved behaviour rather than detecting the change, and pass on
both sides: `spec_position_filepass_is_refused_before_its_payload_is_read`
covers a `FilePass` at record indices 1 through 4 with both payload shapes, and
`encrypted_fixtures_are_refused_without_reading_a_filepass_payload` asserts all
three encrypted fixtures are refused in **exactly two reads, the same as before
the change**, with no payload byte read.

One facade test in `crates/litchi/src/sheet/workbook.rs` pinned the old schedule
as an exact range list. It now asserts the invariants — every globals byte
covered exactly once, no overlapping reads, containment within one window past
the globals end — while leaving its catalog-reuse and metadata assertions
untouched. It was confirmed to **fail** against the pre-change code, because the
old pre-pass read header bytes twice, so it is a strengthening rather than a
weakening.

An independent adversarial review could not construct an input making a fill
unbounded, non-terminating or panicking, and verified that bytes resident behind
a late `FILEPASS` are never framed, hashed, logged or placed in an error
message. Its findings about the plan and survey describing a superseded schedule
are addressed in the amendment recorded in `plan.json`.

## Limitations

No cold-cache, physical-device, remote or range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform result is claimed. The
counters are logical calls over a warm, page-cached, immutable staged copy. The
`xls-tiny` corpora and the nanosecond-scale semantic selectors sit below this
harness clock's useful resolution and should not be read as evidence in either
direction.

Two costs are real and bounded. Byte totals can **rise** on a small workbook: one
fixture reads 4,266 bytes before and 5,225 after while its reads fall from 84 to
12, bounded by the stream length. And an open can read past the globals end when
the first `BoundSheet8` is framed only after the last fill was issued: zero bytes
on the flagship fixture, 93 on the harness locality corpus, at most 3,949 across
the surveyed corpus.

Replay:

```sh
python3 -B -c "import json;print(json.load(open('docs/performance/results/change-0565/latency/analysis.json'))['summary'])"
```
