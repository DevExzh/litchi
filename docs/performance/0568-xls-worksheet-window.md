# 0568: one windowed pass over an XLS worksheet scan

Status: retained. `performance_claim: none` — this record carries deterministic
read and byte counts, and paired latency where the host permitted it.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What this implements

Change [0564](0564-xls-open-read-attribution.md) measured a source-backed
one-cell query at 266 positional reads beyond the open: 177 four-byte record
headers, 88 payload reads and one shared-string fetch, so "a skipped record costs
exactly one read; a consumed record costs exactly two". Change
[0566](0566-xls-worksheet-window-design.md) designed the fix and froze it. This
change implements that design.

`WorksheetScan` keeps its stream cursor — the design warned specifically against
substituting the globals scan's range reader, because worksheet substreams begin
deep in the stream and the cursor does not re-walk the allocation chain — and
replaces the per-record scratch buffer with a retained window. One primitive
fills it: resident bytes cost nothing, otherwise the framed prefix is compacted
away and one cursor read is issued. Fills start at 512 bytes and double to a
64 KiB cap. A skipped payload that extends past the filled end is passed by
seeking, which reads nothing and takes no source observation.

## Why this case is easier than the globals scan

`validate_sheet_offsets` runs at open, before any query. It sorts the sheets by
start offset, rejects duplicates, and gives each sheet an end equal to the next
sheet's start or the stream length. So the upper bound is **exact before the scan
begins**, every fill is clamped to the selected sheet's own validated region, and
there is no spill-and-truncate and no over-read clause of the kind change
[0565](0565-xls-globals-single-pass.md) had to accept.

## Measured effect

Logical calls through a counting source, captured on the same tree before and
after the source change with the tests held constant.

| Selector | reads | bytes | observations |
| --- | ---: | ---: | ---: |
| open (control) | 53 → **53** | 565,201 → **565,201** | 28 → **28** |
| list (control) | 0 → **0** | 0 → **0** | 1 → **1** |
| one cell | 265 → **7** | 1,706 → 37,907 | 268 → 10 |
| one cell, small fixture | 31 → **2** | 150 → 323 | 37 → 8 |
| full text, small fixture | 86 → **4** | 424 → 1,098 | 109 → 27 |

**Open and list are unchanged read-for-read and byte-for-byte on every fixture.**
That is the control, and it is what shows the reduction belongs to the worksheet
scan rather than to anything else.

The reductions are 97.4% on the flagship one-cell query and 93.5% and 95.3% on
the small fixture. **The cost is bytes**: 22.2 times more on the flagship
one-cell query, 2.2 to 2.9 times on the small fixture. That trade is the change,
stated plainly rather than omitted.

The fill schedule taken on the flagship worksheet is
`512, 1024, 2048, 4096, 8192, 16384, 5651` — seven reads totalling 37,907 bytes,
**exactly** what change 0566 predicted, with the last fill clamped by the sheet
boundary and the gaps being skip-seeks. On a 311-byte sheet the first fill clamps
to the whole sheet span, so a small sheet allocates a small buffer.

The design's worked example for a skipped 4,096-byte payload measures two scan
reads and 516 bytes, against the predicted two reads and 516 bytes, where the old
path took 32 reads.

## The density gate, and its untestability on real data

Windowing trades bytes for reads, and the trade inverts on sheets of large
skipped records. Before each fill the scan compares the running mean framed bytes
per record against 1 KiB; above it the fill is exact and the window target resets.

| Synthetic sheet of 200 records with 8,224-byte payloads | reads | bytes |
| --- | ---: | ---: |
| before the change | 231 | 950 |
| after, gate active | **202** | **8,480** |
| after, gate mutated off | 31 | **1,613,508** |

The ungated figure lands within 12 bytes of change 0566's predicted 1,613,496.

**The gate cannot fire on any input in this repository.** A sweep of all 139 XLS
fixtures under `test-data` found none above the threshold; the densest sheet in
the flagship fixture is 544.3 bytes per record, 53% of the trigger, and the
opaque-heavy corpus does not qualify either because its bulk is sibling container
streams rather than records. Its entire justification is a simulated adversarial
input, and that is recorded here as a limitation rather than presented as
coverage.

Lowering the threshold to buy real coverage was considered and **rejected with
numbers**: 512 would put one real sheet back on the exact path, taking it from 7
reads to roughly 200 on a one-cell query, which buys test coverage by removing
the change's benefit. 1 KiB is where a 64 KiB fill stops paying for itself at the
roughly 116 ns per syscall change 0564 derived.

## Behavioural differences

1. **Bytes read rise**, as tabulated above.
2. **A fill may read trailing slack inside the selected sheet's own validated
   region.** One fixture has 1,809 bytes between a worksheet's end record and the
   stream end, of which the scan now reads 249. No byte outside the selected
   sheet's region is ever read, so the existing owner-bounding test survives
   unchanged.
3. **`max_worksheet_scan_bytes` becomes a true read fence**, tighter than the old
   limit-plus-four: with a limit of 2 the scan now reports the limit having read
   **zero** bytes, where it previously read four.
4. **Error precedence changes in two places.** Within one fill, an I/O or
   source-version failure at a later offset precedes a framing, limit or parse
   error of an earlier record in the same fill. And the pre-read byte-limit check
   now runs before that record's oversize, boundary and record-count checks, so a
   configuration with fewer than four bytes of byte budget **and** a record limit
   tripping at the same record reports the byte resource where it reported the
   record resource. No existing test exercises it. The boundary check
   deliberately still runs first, so truncated-tail errors are unchanged.
5. **Source observations fall from one per read to one per fill**, 268 to 10 on
   the flagship one-cell query, with the query-level bracket unchanged.
6. **Cancellation keeps every check site**, so the interval between checks stays
   at one record in CPU terms and grows to one fill only in I/O terms.
7. **Peak retained memory** per scan rises from one payload of at most 8,224
   bytes to at most 64 KiB plus one taken payload, capped by the smaller of the
   window cap, the sheet span and the byte limit.

## Correctness evidence

Two tests were rewritten, because both old contracts are unachievable under any
coalescing scheme — a record's kind is knowable only from its header, and the
header sits inside the window, which is exactly the gap change
[0325](changes/0325-cfb-frame-transaction-rejected.md) named. Both replacements
are themselves change detectors, and both assert the strong form: **no read
begins inside the payload**, and the remainder past one window is never read.

Eight tests were added, seven of them change detectors. The one that passes on
both sides is the opaque-heavy fallback, and that is its purpose: the gate exists
to reproduce today's shape on an adversarial input. It is mutation-checked, as is
the hysteresis test that proves the bound is a running mean and **not a one-way
latch** — exact fills over a dense run, then a window fill returning inside the
sparse run that follows. A further test pins that a real worksheet stays on the
windowed path, and passes with the gate mutated off, which is what makes it
meaningful.

Detector claims were verified by restoring the source file to its committed
content and rerunning, rather than by stashing, so concurrent work was not
disturbed.

`litchi-xls` passes 1,359 tests with zero failures. With change 0570 applied
alongside, the OLE2 crates pass 4,261 tests with zero failures across 154
binaries.

## Paired latency, and what it is actually attributable to

The latency matrix does **not** run on the fixture the counted evidence above
uses. Its ten selectors use the in-memory `xls-comments-opaque-heavy` corpus, and
the candidate binary carried this change **and** change
[0570](0570-cfb-fat-run-batching.md) together. Re-capturing the counted evidence
on that corpus separates them cleanly:

| Selector on the latency corpus | CFB structural reads | selected worksheet reads | total |
| --- | ---: | ---: | ---: |
| `xls_source_backed_open` | 265 → **8** | 0 → 0 | 272 → 15 |
| `xls_source_backed_open_list_worksheets` | 265 → **8** | 0 → 0 | 272 → 15 |
| `xls_source_backed_open_one_cell` | 265 → **8** | 28 → **2** | 300 → 17 |

The structural collapse is change 0570's FAT run batching, at **136,704 bytes
unchanged**, and it appears in all three selectors. This change's component is
the selected-worksheet column, and it is **zero for open and list**. So the
`xls_source_backed_open` and `..._list_worksheets` latency movements below belong
to change 0570 alone; only `..._one_cell` carries both changes. Recording that
plainly matters more than the headline: without the separation this change would
have been credited with gains it did not produce.

Measured on a host verified quiet first — the pinned core at 99.67% to 100% idle
and a one-minute load average of 3.21, reached after waiting 340 seconds — with
40 children at 5 warmups and 60 samples each. 31 of 56 comparisons improve in
both directions.

| Selector | p50, first direction | p50, second direction | attributable to |
| --- | ---: | ---: | --- |
| `xls_source_backed_open_one_cell` | −15.78% | −12.63% | both changes |
| `xls_source_backed_open` | −7.80% | −13.23% | change 0570 |
| `xls_source_backed_open_list_worksheets` | −7.40% | −10.25% | change 0570 |
| three owned-source variants | −6.34% to −8.56% | | change 0570 |

### The noise floor, measured in the same window

An extra same-binary A/A matrix at the same sample count, in the same quiet
window, gives **p50 4.10%, mean 4.60%, p95 7.33%, p99 13.70%**, with zero A/A
review triggers. So the 5% trigger is usable at the median but sits **below** the
p95 and p99 floors.

One comparison is adverse in both directions by more than 5%:
`xls_semantic_one_cell` on `xls-large` at p99, +6.13% and +8.70% — 230.0 ns to
244.1 ns and 250.0 ns, an absolute movement of 14 to 20 nanoseconds. That is
**inside its own statistic's noise floor** of 13.70%, and the same cell's median
swings +44% then −28% within the same run. It is reported rather than excluded,
and it carries no information.

## Limitations

The strict `tools/perf_abba_summary.py` cross-check accepts 4 of the 10
selectors. The six refusals all name the **baseline** leg and reproduce
identically against change 0565's own retained and accepted reports, so they are
pre-existing tool skew rather than a property of this candidate. That is recorded
here as a known gap in the strict checker, not worked around.


No cold-cache, physical-device, remote or range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform result is claimed.

Three coverage gaps are recorded rather than worked around. The density gate has
no real-fixture coverage and cannot acquire any from this corpus. A complete text
extraction of the flagship fixture is **not measurable** through the public text
API: it fails with the same parse error before and after the change, so change
0566's modelled 5,277-read text scan is not reproducible on that fixture, and the
text figures above come from smaller fixtures that complete. And the source-backed
API has no whole-sheet iterator, so a "full cell scan" is N independent one-cell
queries, each re-scanning its sheet; only the per-query ratio is meaningful.
