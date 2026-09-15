# 0584: a fresh OLE2 profile, and the first DOC and PPT measurement in this program

Status: retained, attribution only. `performance_claim: none` — this record
carries per-symbol instruction counts, caller attribution, call counts, chain-link
counts and two independent byte-layout models. **No timing was measured, no
before/after comparison was made, and no optimization is claimed or authorized by
it.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is the post-0579 successor to change [0574](0574-ole2-next-opportunity-survey.md),
taken at `e927e139c` (post-0579, post-0583). It re-ranks the OLE2 read path,
answers a question 0574 could not — because it profiled `open` only — and closes
0574's own statement that "no DOC or PPT scenario was measured at all".

## What was changed

**No production code.** The only files added are this record and its evidence
directory.

## The method, and why it differs from 0574's

Change 0564 established the isolation-pair method: profile a large-sample and a
small-sample child, difference the **program totals**, divide by the sample delta.
This record differences the same pairs **per symbol**. One-time process cost then
cancels exactly rather than approximately — the harness's SHA-256 fixture staging
is 53% of the raw profile and vanishes entirely in the delta, which is what makes
a symbol table of the *operation* rather than of the *process* possible at all.

Three controls tie the result to change 0579's retained evidence:

| check | this profile | 0579 retained |
| --- | ---: | ---: |
| flagship `open` whole-operation Ir | 2,379,097 | 2,380,011 (**0.04%**) |
| flagship `open` FAT chain links | 2,099 | 2,099 (**exact**) |
| flagship `open` reads / bytes | 53 / 565,201 | 53 / 565,201 (**exact**) |

Two byte-identical runs differ by **635 Ir in 678,704,129** — 0.0001%. Maximum
per-symbol drift is 0.47%, confined entirely to glibc malloc internals; every
`litchi_*` symbol drifts below 0.05%. Instruction counts here are effectively
deterministic, which is the property that makes the ranking worth anything.

The binary was built from a detached worktree of `e927e139c` outside the working
copy, because three files were being edited in the working copy while this profile
ran. Its hash is recorded in the packet.

## The worksheet scan is bimodal, and the flagship fixture alone would hide it

Change 0579 left open whether the worksheet scan had become the dominant term. The
answer is that there is no single answer.

| fixture | `open` Ir | `one-cell` Ir | scan delta | scan share of the query |
| --- | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 2,379,097 | 2,616,858 | 237,761 | **9.1%** |
| `54016.xls` | 7,239,535 | 24,525,331 | 17,285,796 | **70.5%** |
| `WithCustomViews.xls` | 915,526 | 937,487 | 21,961 | 2.3% |
| `SimpleWithImages-mac.xls` | 335,079 | 351,604 | 16,526 | 4.7% |

On the flagship the open still dominates the query ten to one. On `54016.xls` the
scan is seven tenths of it. **The cause is not fixture size** — `54016.xls` is
smaller than the flagship. It is that `query_cell` scans to EOF and never stops
when the target cell is found. One query for row 1, column 0 frames **37,929
records**, reads **32,358 payloads**, and parses and processes **29,608 cell
records** to return one value.

That unbounded scan is **out of scope and stays out of scope**: change 0574's
opportunity 6 was rejected on ADR 0005 mandatory-validation grounds, and this
record does not reopen it. It is reported because it sets the multiplier on
everything inside the loop — a per-iteration cost is paid roughly thirty thousand
times per query on this fixture — and because any percentage measured inside that
loop must not be read as a bound on the query.

## Three regimes, not one hot path

`54016.xls` alone splits into two unrelated profiles either side of its open.

**`flagship/open`** — 2,379,097 Ir. Bulk byte movement leads (`memcpy` 28.64%,
`memset` 23.41%), then `SectorChainScratch::collect_exact` 3.94%,
`from_shared_ole_file_with_limits` 3.66%, `read_stream_range_hinted` 2.21%,
`next_chain_sector` 2.12%.

**`54016/open`** — 7,239,535 Ir, of which **49.5% is the SST measure-walk**:
`scan_shared_string_records` 16.08%, `walk_one_shared_string` 11.56%,
`MeasuredText::consume` 9.94%, `SstCursor::read_exact` 8.62%,
`read_formatting_runs` 3.27%.

**`54016/one-cell`** — 24,525,331 Ir, the worksheet-scan regime:
`WorksheetScan::next_frame` 17.32%, `query_cell` 10.67%, `memcpy` 7.80%,
`WorksheetScan::ensure` 7.17%, `process_cell` 6.67%, `read_payload` 6.07%,
`scan_shared_string_records` 4.75%, and then
**`core::ptr::drop_in_place<SourceBackedError>` at 4.20%** — the eighth hottest
symbol of the operation, on a path that returns `Ok`.

Inclusive cost of the entry points:

| | flagship/open | flagship/one-cell | 54016/one-cell |
| --- | ---: | ---: | ---: |
| `SourceBackedWorkbook::from_read_at` | **96.71%** | 87.91% | 29.45% |
| ⤷ `GlobalsBuffer::ensure` | **57.66%** | 52.45% | 6.24% |
| ⤷⤷ `read_stream_range_hinted` | 28.35% | 25.79% | 1.53% |
| ⤷ `parse_globals` (the semantic pass) | 16.58% | 15.03% | 4.33% |
| ⤷ `SharedOleFile::open_with_limits` | 10.42% | 9.47% | 0.65% |
| ⤷ `scan_shared_string_records` | 4.86% | 4.42% | 15.19% |
| `query_cell` | — | 9.02% | **70.47%** |

**Reading the globals bytes costs 3.5× what interpreting them costs** — 57.66%
against 16.58%.

## Chain links are now attributable per caller

0579 could count chain links per open. This record counts them per caller, which
is what makes the next change targetable:

| operation | total | `read_stream_range_hinted` | `stream_cursor_at` | `normalize_state` |
| --- | ---: | ---: | ---: | ---: |
| flagship `open` | 2,099 | 2,099 | — | — |
| flagship `one-cell` | 4,428 | 2,099 | **2,238** | 91 |
| `54016` `one-cell` | 2,920 | 1,110 | 599 | 1,211 |
| `cv` `one-cell` | 629 | 336 | 292 | — |

**`stream_cursor_at` walks more links in one flagship query than the entire
globals scan does**, from two call sites that each restart at the stream's first
sector. 0579's hint reaches neither, because it was applied to
`read_stream_range` only and there is no hinted cursor constructor.

## DOC and PPT, measured for the first time

Change 0574 states plainly that "no DOC or PPT scenario was measured at all". A
throwaway driver was built for this record; its source is retained.

| fixture / op | Ir per op | `memset`+`memcpy` | `litchi-cfb` |
| --- | ---: | ---: | ---: |
| `docbig/open` (1,619,457 B) | 4,962,445 | **54.0%** | 7.2% |
| `docmid/open` (335,360 B) | 3,457,522 | ~26% | 1.6% |
| `pptbig/open` (842,240 B) | 949,401 | **42.2%** | 10.7% |
| `pptmid/open` (385,024 B) | 1,180,448 | **58.1%** | 8.6% |

**DOC cost is not size-driven.** The 65 KB fixture costs 9% *more* per open than
the 1.6 MB one, and `docmid` at a fifth the size of `docbig` costs 70% as much.
Whatever dominates a DOC open is not proportional to the document. The top named
DOC symbol is `litchi_doc::sprm::parse_sprms` at 6.35%; the top named PPT symbol
is `litchi_ppt::records::record::Record::parse_impl` at 4.71%. Neither is large
enough to be the answer, and this record does not have one — it establishes that
the question exists.

Two DOC fixtures were dropped because the library refuses them (`style names and
aliases must be unique`); that refusal is correct behaviour and is recorded so the
corpus is reproducible rather than silently trimmed.

## Two byte-layout models, written independently of the code they predict

To size the `stream_cursor_at` term without trusting the profile's own
attribution, the walk was modelled directly from the CFB and BIFF byte layout in
Python — parsing the FAT, the `Workbook` stream, the `SST` record with its
`Continue` segments, every `LabelSst` index in stream order, and every
`BoundSheet8` start. The model runs no repository code and needs no build.

Shared-string resolves, over every `.xls` fixture under `test-data/`:

| fixture | resolves | ordinals | links today | links resumable | saving |
| --- | ---: | --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 658 | 1,067–1,076 | 703,937 | 143,502 | 79.6% |
| `FormulaEvalTestData.xls` | 1,508 | 31–60 | 66,009 | 10,806 | 83.6% |
| `WithCustomViews.xls` | 25 | 10–136 | 765 | 365 | 52.3% |
| `WithEmbeddedObjects.xls` | 1 | 41 | 41 | 41 | **0.0%** |
| **76 string-bearing fixtures** | | | **3,066,828** | **785,058** | **74.4%** |

Median per-fixture saving is **79.8%**, and **13 of the 76 fixtures save nothing
at all** — single-resolve files with no prior position to resume from. The corpus
total is lower than the median because a few large fixtures carry most of the
weight and have higher backward-step rates, near 20% on both primary fixtures.

Per-sheet cursors, from `BoundSheet8` starts:

| fixture | sheets | ordinals | links today | links resumable | saving |
| --- | ---: | --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 16 | 1,076–2,477 | 28,143 | 2,477 | **91.2%** |
| `HyperlinksOnManySheets.xls` | 3 | 3–6 | 14 | 6 | 57.1% |
| `54016.xls` | 1 | 599 | 599 | 599 | **0.0%** |

The model assumes a resumable position recorded at the cursor's construction
ordinal and never advanced by the cursor's own subsequent reads, and counts FAT
links only. It is a prediction, not a measurement, and change 0585 is where it is
checked against instrumented counts.

## Candidates, each screened against the record set before proposing

| # | candidate | measured share | prior art | risk |
| ---: | --- | --- | --- | --- |
| 1 | resume the cursor's construction walk | 3.36% Ir; **2,238 of 4,428** flagship one-cell chain links | 0579 names it as the open question it does not answer | **low** |
| 2 | keep `SourceBackedError` off the scan success path | 4.20% of `54016` one-cell; 986,154 Ir from `next_frame`, **75,858 drops per query** | no record touches XLS error size; change 0533 is the landed precedent in `litchi-cfb` | **lowest** |
| 3 | index the SST lazily rather than measure-walking it at open | **49.5%** of `54016`'s open | change 0576's own "What this does not do" defers it | high |
| 4 | stop buffering the ~75% of globals bytes never interpreted | ~34% of the flagship open | 0574 opportunity 2, unimplemented | medium-high |
| 5 | stop zero-filling buffers overwritten on the next line | 9–32% of open across XLS, DOC and PPT | 0574 opportunity 4 names **only** the XLS site; the five `litchi-cfb` sites appear in no record | high on policy |

Candidates 1 and 2 are the two this batch attempts. Neither is authorized by this
record; each must clear its own paired measurement, and either may be rejected.

**Candidate 3** moves *when* a malformed SST is refused; change 0576 states it
"needs its own frozen design record" and this record does not supply one.
**Candidate 4** converts one large read into many small ones, which is a
regression on a high-latency source, and its density gate is the same one change
0568 recorded as untestable on any real fixture in this corpus. **Candidate 5** is
blocked by `docs/GOAL.md` rule 10: a safe form needs the reader to *append* into
the buffer so the zero never exists, which is an API change to
`read_stream_range`/`read_chain_into`, not a local edit. Change 0570's warning is
the trap to avoid — it declined a helper that "would have silently zero-filled
where the old path raised a typed error".

### Deliberately not proposed

Early exit from `query_cell` — the largest single number in the profile, 29,608 of
29,608 cell parses wasted on `54016.xls` — is **0574 opportunity 6, rejected** on
ADR 0005 mandatory-validation and `docs/GOAL.md` rule 12 grounds. The tractable
form is a *retained* scan across queries, which 0568's limitations already flag.
`directory_name_data` (7.8% on small fixtures) was tried and **rejected** by change
0554, every primary XLS p50 regressing. `collect_exact` (3.94% flagship) was
rejected by changes 0524, 0548 and 0549, and 0533 forbids reviving 0524. BIFF
double-framing measured 1.78% flagship and 3.21% on `54016`, confirming 0574's
1.51% and remaining too small to carry a change alone.

## Limitations

**Instruction counts rank work, not latency, and this repository has already been
burned by the difference.** Change 0579 removed a dependent-load pointer chase
worth 1.24% of instructions on `54016.xls` and **6.19% of cycles**, and
`GOAL_AUDIT.md` carries the standing instruction that future OLE2 rankings price
cycles where a pointer chase is involved. Candidates 1 and 2 are therefore
*understated* here. Callgrind counts `rep movsb`/`rep stosb` one instruction per
byte, so candidates 4 and 5, which are bulk byte movement, are *overstated* —
change 0574's caveat, unchanged. **No candidate in the table above is authorized
by this record**; each needs its own paired measurement.

No `perf stat`, cycle, cache-miss, branch-miss, wall-clock, allocation, peak-RSS,
cold-cache, range-source, concurrency or cross-platform measurement was taken.
Every figure is single-host, single-toolchain, ASLR-disabled and pinned.

5.0% of the flagship open is `memcpy` called from an address outside the binary's
LOAD segments that callgrind could not name. It is a `memcpy` caller, not a
separate cost, and it is reported rather than silently folded.

The DOC and PPT legs used a throwaway driver rather than the checked harness, so
their operation boundaries are not the ones any registered selector uses. They
establish magnitudes and a shape, not comparable scenario numbers.

The two byte-layout models are predictions derived from file structure. They do
not account for MiniFAT streams, for a hint advanced by the cursor's own reads, or
for any cost other than chain links.

## Retained evidence

[`results/change-0584/`](results/change-0584/README.md) — the consolidated
profile, 199 `callgrind_annotate` outputs across both fixture families, the folded
per-symbol JSON, the caller-attribution trees, both byte-layout models with their
corpus output, the DOC/PPT harness source, and every capture script.
