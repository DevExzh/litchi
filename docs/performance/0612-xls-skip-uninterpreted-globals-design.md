# 0612: the XLS globals skip saves 85% of the open's read bytes and costs 6% to 17% of the open; the density gate fires on real fixtures and the design is frozen unimplemented

Status: retained, design only. `performance_claim: none` — the numbers below are
deterministic counts, callgrind isolation pairs, native `perf stat` medians and
paired wall-clock medians in both directions, reported as evidence rather than
registered as claims. **No file under `crates/` was modified by this change.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was changed

Nothing in production. This record carries four things:

1. the **schedule model** change 0574 opportunity 2 never built — a byte-exact
   replay of today's `parse_globals` fill schedule and of a candidate that skips
   the payloads the semantic pass never interprets, over every `.xls` and `.xlt`
   fixture in the repository, so the trade the opportunity makes (fewer bytes
   for more requests) can be priced per fixture instead of in aggregate;
2. the **frozen design**: the skip primitive, the gate rule and its hysteresis,
   the buffer shape that folds 0574 opportunity 5 into the same change, the
   contracts of change 0565 that survive and the three that do not, the
   read-count trade stated per source kind, and the admission gates;
3. the **measured answer** to change 0587's falsification criterion for XLS-6,
   taken with a measurement scaffold that changes only which stream bytes are
   read and leaves every semantic output byte-identical;
4. the **recommendation not to implement it**, because the fixture the gate
   fires on gets 5.9% slower on an in-memory source and 16.9% slower on a file
   source at p50, against a floor under 1% in the same window.

## The item, and what it asked

Change [0574](0574-ole2-next-opportunity-survey.md) opportunity 2 measured that
75.08% of the globals bytes a source-backed XLS open reads are payloads the
semantic pass never interprets, priced the saving at "27.7 µs of the flagship's
91.8 µs open, 30.2%" from a corpus regression coefficient of 53.4 ns per KiB,
and named the cost as "about +46 requests". Change
[0587](0587-remaining-opportunity-survey.md) ranked it XLS-6 at rank 24 and
wrote the falsification criterion: *falsified if the 103 dense-globals fixtures
regress in cycles more than the sparse ones gain.* It also recorded the one
thing that distinguishes this item from the worksheet density gate of change
[0568](0568-xls-worksheet-window.md), which cannot fire on any input in this
repository: **the globals gate can fire on a real fixture.**

Both halves of that are now measured. The gate fires. The item is still
falsified, for a reason neither record anticipated: the coefficient that sized
it prices a globals byte at about twenty times what moving one actually costs on
this host.

## The mechanism, re-read on this base

`parse_globals` (`crates/litchi-xls/src/workbook/source.rs:1740` at
`1e4198321`) frames every globals record once out of a `GlobalsBuffer` that
holds the stream prefix `[0, filled)`, truncates that buffer to `global_end`,
frames all of it a **second** time through `BiffRecords::with_limits`, and then
interprets a fixed set of record kinds:

| consumer | kinds |
| --- | --- |
| `parse_globals` itself | `BOF` 0x0809, `EOF` 0x000A, `FilePass` 0x002F, `CodePage` 0x0042, `BoundSheet8` 0x0085, `SST` 0x00FC and the `Continue` 0x003C run that follows an `SST` |
| `Formatting::parse_globals` (`number_format/codec.rs`) | `Date1904` 0x0022, `Format` 0x041E, `XF` 0x00E0, `XFCRC` 0x087C, `XFExt` 0x087D, `DXF` 0x088D |

Every other payload is read, framed twice and dropped. The twelve kinds are
confirmed against the tree at the base commit; the model in
[`results/change-0612/model/`](results/change-0612/model/globals_skip_model.py)
cites the constant and its file for each.

## The corpus, modelled per fixture

126 `.xls` and `.xlt` fixtures; three are encrypted and refused inside the exact
prologue, so no schedule exists past it and they are excluded. Over the
remaining 123:

| | bytes | share |
| --- | ---: | ---: |
| globals bytes framed today | 3,969,111 | 100.00% |
| four-byte record headers | 103,880 | 2.62% |
| consumed record payloads | 876,542 | 22.08% |
| **touched closure** | **980,422** | **24.70%** |
| **payloads never interpreted** | **2,988,689** | **75.30%** |

This reproduces change 0574's closure figures — 75.08% there against 75.30%
here, the difference being the three encrypted fixtures 0574 counted and this
model does not — from a script that shares no code with it, and the flagship's
composition to the byte: 551,377 globals bytes over 621 records, 20,417 of
closure (3.70%), 530,960 skippable, of which 508,135 are `Continue` payloads and
**every one of those continues the single `MsoDrawingGroup` record**, verified by
attributing each `Continue` run to the record it continues.

**One correction to the survey.** Change 0587 describes that chain as "524,839
bytes over about 261 records, about 2 KiB per record, against a 1 KiB
threshold". It is not 261 records of 2 KiB. It is **64 payloads at or above
1 KiB holding 524,583 bytes**, almost all of them at the 8,224-byte BIFF maximum:
one `MsoDrawingGroup` of 16,448 bytes in two frames plus about 62 maximal
`Continue` records. That matters, because a 2 KiB skippable payload is *below*
the break-even against one request on a file source and an 8 KiB one is four
times above it. The item was ranked on a per-record size that is four times too
small, in the direction that made it look marginal when the byte arithmetic is
in fact favourable — and it still loses.

### Concentration, and how many fixtures the gate can reach

**90% of all skippable bytes live in 13 of the 123 fixtures**, and 101 of 123
consume more than half their globals. The reachable set is smaller still,
because a gate has to find a payload large enough to be worth a request:

| largest skipped payload on the fixture | fixtures |
| --- | ---: |
| ≥ 1 KiB | 25 |
| ≥ 4 KiB | 13 |
| ≥ 8 KiB | 12 |

## The design, frozen

### The skip primitive

The worksheet scan already has it. `WorksheetScan::skip_payload`
(`source.rs:2454`) passes a payload that ends past the filled end with
`SharedOleStreamCursor::skip_forward`, which publishes no bytes and — documented
at `litchi-cfb/src/shared.rs:2757` — takes **no source-version fence**; a
payload already resident costs nothing and is simply stepped over. The globals
scan reads through `SharedOleFile::read_stream_range_hinted`, which has no
skip, so the design needs one of two shapes: a retained cursor
(`stream_cursor_at_hinted` once, then `read_exact`/`skip_forward`), which is
change 0574 opportunity 3 folded in, or a range reader called at the next
offset it actually wants, leaving the intervening bytes unfetched. Both were
built and measured. Both lose; see *Measured*.

### The gate rule, and why 0568's rule cannot be ported

Change 0568's gate is a **cumulative running mean** of framed bytes per record,
recomputed before every fill, above which the fill becomes exact: "The mean is
recomputed for every fill, so it is its own hysteresis." Ported unchanged to the
globals scan it is **inert**. Modelled over the corpus it saves 3.0% of the
bytes read, against 59.0% for the rule below, and it fires on **zero** of 123
fixtures. Two reasons, and both are properties of the data rather than of the
threshold:

- The skippable mass is one dense run inside an otherwise sparse stream. The
  flagship's cumulative mean skipped bytes per record is 530,960/621 = 855,
  below a 1 KiB threshold, even though the run itself is 8 KiB per record.
- 0568's scan had to decide a fill size *before* framing the record the fill
  would cover, so a predictive rule was the only rule available. The globals
  skip decision is taken *after* the header is framed, so the payload length is
  known exactly. The gate can be exact where 0568's had to be a forecast.

The rule this design freezes is therefore two parts:

1. **Per-record, exact.** After framing a header, seek past the payload only
   when the kind is not consumed *and* `payload_len >= GLOBALS_SKIP_MIN_PAYLOAD`
   and the payload ends past the filled end. A shorter skippable payload is read
   through, because a request costs more than its bytes.
2. **Hysteresis in the schedule, not in a counter.** A seek resets the window
   target to `GLOBALS_FIRST_WINDOW_BYTES`. Without that reset the target keeps
   doubling to 64 KiB and the next fill reads straight back over the next
   skippable payload, so the skip fires at most once per window: modelled, that
   costs 55 of the 59 percentage points the rule saves. The target doubles again
   as soon as consumed records resume, so a dense run returns to windowed fills
   by itself. No state is carried; the schedule remembers.

`GLOBALS_SKIP_MIN_PAYLOAD` at 1 KiB matches `WORKSHEET_DENSE_FRAME_BYTES` and is
the value measured below. 2 KiB is the break-even at change 0564's 116 ns per
request and change 0574's 53.4 ns per KiB; the corpus is insensitive between
1 KiB and 2 KiB because no fixture has a skippable payload in that band, and
8 KiB costs 51,438 bytes of the saving to remove the last two net-worse
fixtures.

### The buffer shape, folding 0574 opportunity 5

Once payloads are skipped the retained buffer is no longer the stream prefix, so
the second framing (`BiffRecords::with_limits(&bytes, …)`, opportunity 5, 1.51%
of the flagship open) cannot run over it unchanged and the two changes are one
change, exactly as 0574 said. The shape:

- `GlobalsBuffer::bytes` becomes a **compacted** buffer holding, contiguously,
  the header and payload of every record whose payload is consumed, and
  **nothing at all** for a record whose payload is skipped. Retaining a skipped
  record's header alone is not an option: a header declares a payload length,
  so a header without its payload does not frame, and the skipped record must be
  absent from the compacted buffer entirely.
- Dropping them is invisible to both consumers. `Formatting::parse_globals`
  matches only consumed kinds and every counter it keeps (`MAX_FORMAT_RECORDS`,
  the XF index, the duplicate-format map) is relative to consumed records.
  `parse_globals`' own match is the same. The `first`/`last` `BOF`/`EOF` checks
  hold because both kinds are consumed and keep their positions. The `SST`
  `Continue` run is preserved exactly, because a `Continue` is consumed if and
  only if the run it belongs to follows an `SST`, and that is decided in the
  first pass where the preceding non-`Continue` kind is known.
- The second framing then runs over 20,417 bytes and about 350 records on the
  flagship instead of 551,377 bytes and 621, which is opportunity 5 taken for
  free. Every limit the second framing enforces is already enforced by the first
  pass: `max_records` by the `max_global_records` check, `max_record_bytes` by
  the `MAX_RECORD_BYTES` check, `max_input_bytes` by the `max_global_bytes`
  check against a buffer that is now strictly smaller.

**The design uncovers one contract the survey did not name.** `RecordRef`
carries `offset`, the header's offset *in the source stream*, and
`scan_shared_string_records` (`records.rs:1315`) writes it into
`SharedStringSstSegment::source_offset`. That value is not diagnostic: the
source-backed resolve path (`source.rs:3261-3266`) re-opens a CFB cursor and
calls `skip_to(source_offset)` to fetch a shared string's bytes at query time.
A `RecordRef` framed out of a compacted buffer carries the **compacted** offset,
so every shared string on every fixture would be read from the wrong stream
position. The design must therefore carry the stream offset of each retained
record alongside the compacted buffer — one `u64` per retained record, or one
per `SST` frame — and `scan_shared_string_records` must take it from there
rather than from `RecordRef::offset()`. This is not optional and it is not
small: it is the single place where the compacted shape is observable outside
the globals pass.

### Contracts of change 0565 that survive, and the three that do not

Surviving, unchanged:

- **Bytes buffered and never published.** A `FilePass` framed after the exact
  prologue may still have payload bytes resident in a fill; they are never
  framed, interpreted, logged, put in an error message or published. Skipping
  makes this strictly stronger, since fewer bytes are ever resident.
- **The exact prologue.** The first four records, and by read coupling the
  fifth, are fetched exactly, so an encrypted workbook is still refused before
  any payload byte is requested. The design applies the gate only outside the
  prologue.
- **The `BoundSheet8` fill clamp** and its running-minimum semantics, and the
  `max_global_bytes` plus-four bound: the four header bytes that prove a record
  crosses the limit are read, and the limit is reported with the record's real
  end as `observed`.
- **`max_global_bytes` as a bound on the globals span.** The check is against
  the stream offset `end`, which does not move, so the typed error and its
  `observed` value are identical. What changes is its *meaning*: it stops being
  an upper bound on bytes read and becomes a strictly looser bound on bytes
  retained.

Changed, and each needs a test rewritten rather than deleted:

1. **"Every globals byte is read exactly once, over-read exactly zero."**
   `source_backed_open_reads_each_global_byte_once` and the facade test in
   `crates/litchi/src/sheet/workbook.rs` assert the covering property over
   `[0, global_end)`. Under the skip it becomes: every *retained* globals byte
   is read exactly once, no two reads overlap, and no read begins inside a
   skipped payload. Both replacements are still change detectors and both are
   still stronger than the pre-0565 shape.
2. **A read failure confined to a skipped span no longer surfaces at open.**
   The allocation chain is still walked past those sectors, whether by
   `skip_forward` or by the range reader resolving a later offset, so chain
   corruption, ownership violations and length inconsistencies are detected
   exactly as before. What is no longer detected at open is an I/O error or a
   `SourceChanged` whose only evidence lies in bytes nothing interprets.
3. **Error precedence moves back in one place.** Change 0565 accepted that a
   read failure at a later offset can pre-empt a `FilePass` or limit error of an
   earlier record in the same fill. The skip removes the read that carried that
   precedence, so a limit error of a record inside a skipped span is now
   reported where an I/O error previously could have pre-empted it. That is the
   *opposite* direction from 0565's accepted change and closer to the pre-0565
   order; it is still a move, and it is still untested by any existing fixture.

### The read-count trade, per source kind

| source | what one extra read costs | what one skipped KiB saves | verdict |
| --- | --- | --- | --- |
| owned in-memory (`ReadAt` over a slice) | a bounds check, a `find_entry` path resolution and a `memcpy` of the run — **measured** 5.54% of the flagship open for 48 of them | a `memcpy` of a KiB out of the page cache, bandwidth-bound | **measured net loss**, +5.54% cycles |
| file-backed (`FileSource`, what `from_path` builds) | the above plus one `pread64` and one `statx` freshness observation — **measured** 17.22% of the flagship open for 48 reads and 56 observations | the same `memcpy`, plus a page-cache lookup | **measured net loss**, +17.22% cycles |
| range or remote | one round trip, unbounded in this program's measurements | one KiB not transferred | **unmeasured.** This is the only source kind where the trade could invert, and no such source exists in this repository |

The freshness observations are the part the survey did not count: they rise from
29 to 85 on a flagship open, one per read, and on a `from_path` workbook each is
an `fstat`. That is the same per-observation cost item XLS-4 of change 0587
names, arriving here as a cost rather than a saving.

### Admission gates the design would have to pass

Frozen so a later batch does not re-derive them:

- **Counters, every fixture, every scenario.** Reads, read bytes and source
  observations for `open`, `list` and `one-cell` on all 123 modellable
  fixtures, against the base. The 104 fixtures the gate does not reach must be
  identical read-for-read and byte-for-byte; the 19 it reaches must match the
  model's predicted schedule exactly.
- **A differential over the whole corpus.** `ParsedGlobals` — sheet entries,
  `Formatting`, the SST segment and entry locators, the encoding — compared
  between the skipping and the non-skipping scan on every fixture. This is the
  only gate that catches a consumed kind omitted from the predicate, which is
  the standing maintenance hazard of the design: the predicate and the two match
  statements it mirrors are in different files.
- **Error identity on the refused fixtures.** The three encrypted fixtures must
  still be refused in exactly two reads with no `FilePass` payload byte read, and
  the `max_global_bytes`, `max_global_records` and truncated-tail errors must
  keep their variant, `resource`, `observed` and `maximum`.
- **Paired timing on both dense and sparse globals**, in both directions, with an
  A/A floor in the same window, on both `owned-readat` and `file-source`, and on
  at least the flagship (the gate fires, sparse globals), `54016.xls` (the gate
  never fires, 3,421 globals records, so the per-record gate test is the whole
  cost) and `WithCustomViews.xls` (the gate never fires, 98.09% of its globals
  consumed).

Every one of those was run here. The last one is why the design is not
implemented.

## Measured

Host AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0,
valgrind 3.26.0, perf 7.0.14; CPU 8, `taskset` pinned under `setarch x86_64 -R`
for every counted and timed run; seven other measurement agents were active on
the host throughout, which is what the floors are for. Legs, all
`tools/perf-baseline`'s `xls_source_attribution` built `--release --locked` from
the same worktree at `1e4198321`:

| leg | what it is |
| --- | --- |
| `base` | the unmodified base |
| `skip` | **measurement scaffold, not a candidate.** The gate at 1 KiB, applied to which stream bytes are fetched. The buffer keeps today's shape — one contiguous `[0, global_end)` allocation, zero-filled in full — so allocation, zero-fill, both framings and every semantic output are byte-identical to the base, and the skipped bytes are simply zeros nothing reads |
| `cursor` | **measurement scaffold.** Change 0574 opportunity 3 alone: one retained `SharedOleStreamCursor` in place of one `read_stream_range_hinted` per fill, no skip |
| `skipcur` | both together — the design in its best form |
| `aa`, `bb` | byte-for-byte copies of `base`, the floor |

The `skip` scaffold is a **lower bound on the benefit and an exact price of the
cost**: it takes none of the three further savings a real candidate would take
(the zero-fill of the skipped bytes, the smaller retained allocation, the second
framing of the skipped records) and pays nothing a real candidate would pay
beyond them (the compaction, the stream-offset table, the partial-payload
truncate).

### Deterministic counters, the control

`file-source` mode, every counter identical across the five samples of every
child; `owned-readat` gives the same numbers.

| fixture | operation | reads | read bytes | observations |
| --- | --- | ---: | ---: | ---: |
| flagship | open / list | 53 → **101** | 565,201 → **82,867** (−85.34%) | 29 → **85** |
| flagship | one cell | 61 → **109** | 603,115 → **120,781** (−79.97%) | 42 → **98** |
| `WithCustomViews.xls` | open / list | 16 → **16** | 110,242 → **110,242** | 22 → **22** |
| `WithCustomViews.xls` | one cell | 17 → **17** | 110,656 → **110,656** | 25 → **25** |
| `54016.xls` | open / list | 40 → **40** | 317,171 → **317,171** | 25 → **25** |
| `54016.xls` | one cell | 65 → **65** | 932,993 → **932,993** | 43 → **43** |

The two fixtures the gate does not reach are **identical on every counter, every
operation, both source modes**. That is the control, and it is what makes the
timing on them attributable to the gate test in the frame loop and to nothing
else. The flagship's 565,201 → 82,867 matches the model's predicted globals
saving of 482,334 bytes exactly.

The model's own reads and the implementation's differ, and the reason is
recorded rather than reconciled away: the model counts logical fills (20 → 76 on
the flagship) and the harness counts positional reads at the source, which is
where a logical range splits into contiguous sector runs (53 → 101). The
deltas, +56 and +48, differ because smaller ranges split less.

### Callgrind, instructions per operation, owned source

Isolation pairs at small and large sample counts, differenced and divided by the
extra operations.

| fixture | op | base | skip | base − skip |
| --- | --- | ---: | ---: | ---: |
| flagship | open | 2,328,670 | 1,971,708 | **356,962 (15.33%)** |
| flagship | one cell | 2,551,158 | 2,202,623 | 348,535 (13.66%) |
| `WithCustomViews.xls` | open | 809,190 | 812,997 | **−3,807 (−0.47%)** |
| `54016.xls` | open | 5,315,124 | 5,372,720 | **−57,596 (−1.08%)** |

**This table is the trap, and the next one is why.** 412,826 of the flagship's
356,962-instruction saving is `__memcpy_avx_unaligned_erms`, which callgrind
prices at one instruction per byte. Natively the same bytes move in a handful of
retired instructions.

The two fixtures the gate never reaches get *slower* by exactly the gate test:
the whole of `54016.xls`'s +57,596 and `WithCustomViews.xls`'s +3,807 lands in
`parse_globals` (inlined into `SourceBackedWorkbook::from_sha…`), and it is
**16.8 and 17.5 instructions per globals record** — the same number twice, over
3,421 and 218 records. A per-record kind test in the globals frame loop costs
17 instructions per record whether or not it ever fires.

### Native counters, the measurement that decides it

`perf stat` isolation pairs, 100 against 1,100 samples, median of five
repetitions per leg, per open.

| fixture | mode | base cycles | skip | skipcur | cursor |
| --- | --- | ---: | ---: | ---: | ---: |
| flagship | owned | 348,468 | 367,770 **+5.54%** | 403,425 **+15.77%** | 386,455 **+10.90%** |
| flagship | file | 438,220 | 513,667 **+17.22%** | 541,401 **+23.55%** | 479,271 **+9.37%** |
| `WithCustomViews.xls` | owned | 129,806 | 127,982 −1.41% | 133,234 +2.64% | 134,613 +3.70% |
| `WithCustomViews.xls` | file | 171,631 | 173,637 +1.17% | 172,650 +0.59% | 173,077 +0.84% |
| `54016.xls` | owned | 875,353 | 886,157 **+1.23%** | 906,683 +3.58% | 901,029 +2.93% |
| `54016.xls` | file | 948,237 | 961,462 **+1.39%** | 979,092 +3.25% | 967,366 +2.02% |

Native **instructions** on the flagship open go the other way from callgrind:
1,179,292 → 1,284,707, **+8.94%**. The two tables reconcile once the per-byte
accounting of `rep movsb` and `rep stosb` is taken out, which is what changes
0574 and 0604 both warned about: callgrind's delta outside `memcpy` and
`memset` is **+84,374** instructions, against **+105,415** measured natively.
The 21,041 that remain are the native instructions the string operations
themselves retire plus the inlining differences between the two builds; the
sign, and four fifths of the magnitude, are accounted for.

**A/A floor, same window, same metric.** `aa − bb` is +0.95%, +0.34% and −0.61%
in cycles and −0.00%, −0.13%, +0.02% in instructions; `base − aa` is +0.73%,
+0.00%, +0.44%. So the floor on this metric is under 1%, and +5.54%, +17.22% and
+1.23% are all outside it.

### Where the instructions go, flagship open, base → skip

Self cost per open, callgrind, the sites that move by more than 1,500:

| site | delta | what it is |
| --- | ---: | --- |
| `__memcpy_avx_unaligned_erms` | −412,826 | the byte saving, priced per byte |
| `__memset_avx2_unaligned_erms` | −28,510 | incidental; the scaffold still zero-fills the holes |
| `next_chain_sector` | −21,336 | the reader walks the chain to copy each sector, and copies fewer |
| `read_stream_range_hinted` self | −16,508 | |
| `directory_name_data` | **+33,656** | `read_stream_range_hinted` resolves the directory entry **by path on every call** |
| `find_entry` | **+16,016** | the same |
| `realloc` + `_int_realloc` + `finish_grow` | +20,055 | 48 more buffer extensions |
| `parse_globals` (inlined) | +10,514 | the gate test, 621 records |
| `SmallVec` extend | +9,352 | the path vector each `find_entry` builds |
| `GlobalsBuffer::ensure` | +5,496 | 48 more calls |
| `Vec::resize`, `try_reserve_exact`, `ReadAt::read_exact_at` | +7,896 | |
| `clock_gettime`, `Instant::elapsed`, `Timespec::{now,sub}` | +12,168 | **harness instrumentation**, not production: `xls_source_attribution` times every read call |

Two of those are worth naming. The harness's own per-read timers are about 14%
of the non-`memcpy` cost and are not a property of the change; correcting for
them at a native cost of roughly 30 cycles per timestamp moves the flagship's
+19,302-cycle deficit to about +16,400. And `directory_name_data` plus
`find_entry` plus the `SmallVec` is +59,024 instructions spent re-resolving one
directory entry 48 extra times — which is what made the retained cursor worth
measuring, because a cursor resolves it once.

### The retained cursor is slower, and that is its own finding

Change 0574 opportunity 3 — give the globals scan a retained stream cursor —
was ranked, never measured and never implemented. Measured here in isolation,
on this base, it **regresses**: +10.90% and +9.37% of the flagship open in
cycles, +2.93% and +2.02% on `54016.xls`. The callgrind account is unambiguous:
the cursor removes `read_stream_range_hinted` (−52,567), halves
`next_chain_sector` (50,376 → 25,824) and removes most of the directory
resolution (−21,450), for −97,145; and it adds
`SharedOleStreamCursor::{read_exact, normalize_state, physical_span,
advance_within}` for +150,272, a net +47,210 instructions per flagship open. The
per-fill state machine costs more than the path lookup it saves on a scan that
issues about twenty fills. Opportunity 3's premise — that the globals scan pays
a quadratic prefix walk — was true when 0574 wrote it and is no longer true on
this base, because `read_stream_range_hinted`'s chain hint already resolves it.

This is scoped to the one cursor shape built here (`stream_cursor_at_hinted`
once, then `read_exact` per fill). It is not a claim that no cursor could win.

### Paired wall clock, both directions, with the floor

`A1 B1 B2 A2`, 400 samples per round after 50 warm-ups, p50 in nanoseconds;
`A2/A1` is the same binary against itself in the same window. Full p50, mean,
p95 and p99 for all 96 comparisons are in
[`analysis.txt`](results/change-0612/analysis.txt).

| B leg | fixture | mode | A1 | B1 | B2 | A2 | B1/A1 | B2/A2 | floor A2/A1 |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `skip` | flagship | owned | 72,810 | 77,110 | 77,230 | 72,550 | **+5.91%** | **+6.45%** | −0.36% |
| `skip` | flagship | file | 91,031 | 106,441 | 106,461 | 91,500 | **+16.93%** | **+16.35%** | +0.52% |
| `skip` | `WithCustomViews` | owned | 27,820 | 27,120 | 27,220 | 27,791 | −2.52% | −2.05% | −0.10% |
| `skip` | `WithCustomViews` | file | 36,711 | 34,980 | 35,000 | 35,540 | −4.72% | −1.52% | **−3.19%** |
| `skip` | `54016` | owned | 193,891 | 195,471 | 195,241 | 193,461 | +0.81% | +0.92% | −0.22% |
| `skip` | `54016` | file | 212,081 | 214,002 | 215,271 | 211,521 | +0.91% | +1.77% | −0.26% |
| `skipcur` | flagship | owned | 71,930 | 85,370 | 83,980 | 72,400 | **+18.68%** | **+15.99%** | +0.65% |
| `skipcur` | flagship | file | 90,201 | 111,971 | 113,741 | 90,451 | **+24.13%** | **+25.75%** | +0.28% |
| `skipcur` | `54016` | owned | 192,471 | 200,871 | 201,370 | 194,531 | +4.36% | +3.52% | +1.07% |
| `skipcur` | `54016` | file | 209,651 | 217,721 | 219,231 | 210,081 | +3.85% | +4.36% | +0.21% |

**The A/A matrix**, a byte-identical copy of the base as the B leg, run through
the identical pipeline: p50 −0.98%, −0.46%, +0.47%, −0.62%, −0.60%, −0.13% in
the first direction, with self-floors from −0.59% to +0.11%. The window was
quiet by this host's standards — the one-minute load average fell from 29 to 7.5
across the timing runs — so the floor here is **under 1% at p50**, not the 4%
change 0568 measured, and the regressions above are five to twenty-five times
it.

`WithCustomViews.xls` is the one cell where the `skip` leg is consistently
*faster*, by 2.05% to 2.52% on a fixture whose counters are identical to the
byte. Its file-source pair is inside that cell's own −3.19% floor. The
owned-source pair is not, and it has no mechanism: it is code layout, and it is
reported rather than explained away. The `skipcur` leg on the same cell is
+2.32%/+2.06%, which is the opposite sign from the same fixture under `skip`,
and that pair of results together is the size of what layout can do here.

### The coefficient that sized the item

Change 0574 priced this opportunity at 53.4 ns per KiB of skipped globals
payload, derived from a corpus regression of open time against SST and globals
byte volume, giving 27.7 µs on the flagship and 30.2% of its open. The
prediction for the schedule measured here is explicit: 482,334 skipped bytes
worth **−25.2 µs**, plus 48 extra requests at change 0564's 116 ns worth
**+5.6 µs**, a net **−19.6 µs** on a 72.8 µs operation. Measured, the flagship
open moves from 72,810 ns to **77,110 ns, +4,300 ns** on an owned source and
from 91,031 ns to 106,441 ns, **+15,410 ns**, on a file source.

The byte term is the half that is wrong, and change 0604 already holds the
instrument to say by how much. It measured this exact buffer — the XLS globals
zero-fill — at 9,339 cycles for 565,713 bytes, **0.0165 cycles per byte**, and
its copying sibling at 9,307 for the same bytes. At that rate 482,334 bytes are
about 7,960 cycles, **1.7 µs** at the 4.79 GHz this open runs at: modelled, not
measured here, and about **fifteen times smaller** than the coefficient's 25.2 µs.
A regression across fixtures attributes to bytes everything that scales with them
— container geometry, record counts, SST size, allocation, chain length — and
none of that is removed by not reading the bytes. The same correction applies to
XLS-7, which change 0587 sizes at "23.0% of the flagship open" from the same
family of upper bounds and which change 0604 has already re-priced natively at
2.94%.

## Correctness evidence

No production code changed, so there is no new test. What was verified:

- **The scaffold produces identical semantic output.** `xls_source_attribution`
  builds an independent eager-parser oracle per run; the worksheet count and
  names (16, 3, 1) and the selected cell (`string:4:Date`,
  `float:4024000000000000`, none) are identical across `base`, `skip` and
  `skipcur` on all three fixtures. That is what makes the scaffold a control
  rather than a candidate: it changes which bytes are fetched and nothing else.
- **The model is an independent oracle for the counters.** It shares no code
  with `litchi-xls` — a pure-arithmetic CFB and BIFF8 walk over `olefile` — and
  predicts the flagship's byte saving (482,334) to the byte, the two unchanged
  fixtures as unchanged, and change 0574's retained closure figures.
- **Gates, on the clean worktree at the base commit**, in
  [`gates.txt`](results/change-0612/gates.txt): `cargo fmt --all --check` clean;
  `cargo clippy -p litchi-xls --all-targets` clean; `cargo test -p litchi-xls`
  **72 binaries, 1,390 passed, 0 failed, 1 ignored**; `cargo doc -p litchi-xls
  --no-deps` clean.
- **`git diff 1e4198321 -- crates/` is empty**, verified before the gates and
  before the commit; both scaffolds were applied, built, measured and reverted,
  and their patches are retained in the packet.

## Validation preserved

No validation moved, because nothing was implemented. The design's own reading,
recorded so a later batch does not have to re-derive it: **no skipped payload is
validated today.** Both consumers reach a skipped kind through a `_ => {}` arm,
so skipping removes no check. What skipping removes is the *fetch*, and with it
the open-time detection of an I/O failure confined to sectors nothing
interprets; the allocation chain is still walked over those sectors, so every
structural defence of ADR 0006 — ownership, overlap, chain consistency, length —
is untouched. `max_global_bytes`, `max_global_records` and
`litchi_biff::MAX_RECORD_BYTES` keep their check order, their typed variants and
their `observed` values.

## Limitations

- **Not claimed:** any speedup, any cold-cache, physical-device, **range-source**,
  peak-RSS, allocation-count, concurrency-scaling, real-producer or
  cross-platform result. The range-source exclusion is the one that matters: it
  is the only source kind on which this design could win, and this repository
  has none, so its trade is modelled from 116 ns per request and never measured.
- **The `skip` leg is a lower bound on the benefit**, not the change. A real
  candidate would also save the zero-fill of the skipped bytes (from change
  0604's measured 0.0165 cycles per byte, about 8,760 cycles on the flagship),
  the smaller retained allocation, and the second framing of 530,960 bytes; and
  it would pay for the compaction, the per-record stream-offset table the
  `source_offset` contract forces, and the partial-payload truncate. Neither
  side is measured, because nothing was implemented. The measured deficit is
  19,302 cycles owned and 75,447 file-source; the unmeasured savings are of the
  order of the first and nowhere near the second.
- **The gate threshold is measured at one value.** 1 KiB and 2 KiB are
  indistinguishable on this corpus because no fixture carries a skippable
  payload between them; 4 KiB and 8 KiB are modelled but not timed.
- **Only `open`, `list` and `one-cell` exist as selectors.** XLS full text and
  all-cells cannot be timed at all, which is change 0587's standing XLS
  measurement blocker. Both would make the flagship's regression larger, not
  smaller, because both re-enter the same open.
- **Three fixtures are excluded** — `password.xls`, `35897-type4.xls`,
  `xor-encryption-abc.xls` — because they carry a `FilePass` and are refused
  inside the exact prologue.
- **The 19 fixtures the gate reaches are 15 distinct globals shapes**; four are
  byte-identical copies of two originals.
- **The cursor result is one implementation.** A cursor whose per-fill state
  machine were cheaper than `normalize_state` + `physical_span` +
  `advance_within` might change it.

## What this record does not decide

XLS-7 (the globals zero-fill) is not decided here, but its sizing shares the
coefficient corrected above and change 0604 has already declined its CFB
sibling. 0574 opportunity 5 (frame the globals once) is **not** falsified by this
record: it is entangled with the buffer shape only because the *skip* changes
that shape, and on its own it is 1.51% of the flagship open with none of this
change's read cost. It is left open and is the only part of XLS-6 worth taking
forward.

## Retained evidence

[`results/change-0612/README.md`](results/change-0612/README.md) — the model and
its folder, both scaffold patches, the counters, the callgrind and `perf`
isolation pairs for four legs, the folded latency summary, the gates and the
decision record.

Replay:

```sh
cd docs/performance/results/change-0612
python3 model/globals_skip_model.py --repo /home/zhuhe/code/litchi --out model/globals-skip.json
python3 model/fold_model.py
python3 scripts/analyze.py callgrind perf perf-file
python3 scripts/fold_latency.py --summary latency-summary.json
```
