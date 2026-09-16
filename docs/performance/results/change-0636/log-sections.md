# Log paragraphs for change 0636

Four blocks, one per merged log. Each is written to be prepended as the newest
section of its file.

## For `HOTSPOTS.md`

## 0636 — the XLS range-source hotspot is validation, not the whole-sheet walk

Change 0630's item 18 aimed a cursor read-ahead at change 0627's most expensive
range-source scenario, the `54016.xls` whole-sheet walk: 16,145 requests, 16,061
of them at most 512 bytes. Attribution says the target is wrong twice over.
First, the walk's small requests are not the worksheet frame loop — change 0568
already windows that, and the probe sees its exact `512, 1024, 2048, …` schedule
— they are **shared-string resolutions**, one fresh `stream_cursor_at_hinted`
cursor and one read per string, 6,214 of them 3 bytes long, jumping backwards
4,444 times; a read-ahead has no second read on those cursors to serve. Second,
the header-then-payload loop the item describes is `StreamingRecordReader` in
`crates/litchi-xls/src/validation.rs`, and it is **five times more expensive**
than the walk on the same fixture: `litchi_xls::validation::validate_source`,
which takes an `Arc<dyn ReadAt>` and is therefore a range-source entry point,
costs **82,727 positional reads and 82,705 source observations** on
`54016.xls`, 41,358 of those reads exactly four bytes. A bounded window over the
cursor takes it to **60 reads and 31 observations** for the same 937,418 bytes;
across 113 XLS fixtures validation falls from 264,622 reads to 1,842 and from
275,925 observations to 1,699 for 35 more bytes in total, with every one of 565
frozen outcome digests identical. Natively that is −73.84% cycles on the owned
source and −98.31% p50 over `FileSource`, where 82,727 `pread` and 82,705
`fstat` calls become 60 and 31. The walk itself is untouched, and the
shared-string cursor is now the named next item on that axis.
[Change and limitations](0636-cfb-cursor-bounded-window.md);
[evidence](results/change-0636/README.md).

## For `GOAL_AUDIT.md`

## 0636 — a range-source entry point priced and fixed, and a queue item re-aimed

`docs/GOAL.md` names caller-supplied remote and range sources as a benchmarked
dimension, and change 0627 opened that axis for OLE2. This change is the first
production change measured on it, and it moves a P1 row rather than closing one:
`litchi_xls::validation::validate_source` takes an `Arc<dyn ReadAt>` and was
costing 82,727 positional requests on a 984,576-byte workbook — modelled at 82.7
seconds of fixed service at change 0627's transport — and now costs 60,
modelled at 69 ms. It also corrects the audit's working picture of where the XLS
range-source cost sits: the whole-sheet walk 0627 flagged is dominated by
per-shared-string cursor constructions at random offsets, not by the framing
loop, so the "finish source-backed CRUD adoption across formats" row still owns
that work and it is a different mechanism. What this change does **not** supply
is an observed range-source measurement for validation: no selector exists for
it, and building one before the change would have cost 82.7 s per sample, so the
counts carry the argument and the 0627 model only prices it. The read path is
unchanged on every axis measured — identical reads, bytes and observations over
113 fixtures, and identical request-sequence digests over 0627's five
range-source selectors — so no existing coverage row changes status. OLE2 and
OOXML remain the active priority; ODF stays deferred and iWork excluded.
[Change and limitations](0636-cfb-cursor-bounded-window.md);
[evidence](results/change-0636/README.md).

## For `REPORT.md`

## 0636 — a bounded window over the CFB stream cursor, priced on two source kinds

`litchi_cfb::BufferedOleStreamCursor` wraps one `SharedOleStreamCursor` with a
window of at most 64 KiB, filled on change 0568's `512 → 64 KiB` doubling
schedule through the cursor's own validated chain walk, and the XLS validation
record walk is its one caller. Deterministic counts over every `.xls` fixture
under `test-data` (126 files, 113 of which open, 565 operation rows per leg):
validation falls from **264,622 positional reads to 1,842** and from **275,925
source observations to 1,699**, for **35 more bytes in total** and a maximum
per-fixture over-read of **3 bytes**; `open`, `list`, `all-cells` and
`full-text` are identical read for read, byte for byte and observation for
observation, and all 565 frozen outcome digests match. Change 0627's five XLS
range-source selectors reproduce identical request counts, bytes and complete
ordered request-sequence SHA-256 on both legs. `perf stat` isolation pairs put
validation at **−73.84% cycles / −63.29% instructions** on `54016.xls`,
**−65.68% / −58.26%** on `WithCustomViews.xls` and **−43.23% / −45.85%** on
`ConditionalFormattingSamples.xls`, with the read-path controls at **+0.00%**
instructions. Paired timing, 120 samples per leg, order A1 B1 B2 A2 on CPU 12:
**−98.31% p50** validating `54016.xls` over `litchi_core::FileSource` (27.96 ms
→ 0.47 ms), −94.12% on `ConditionalFormattingSamples.xls`, and −73.89%, −66.45%
and −41.89% on the owned-source legs, against A/A floors of at most 1.20% at
p50. Two read-path rows moved the wrong way in an earlier, busier window —
`all-cells` +2.88% and `full-text` +6.92% p50 on `54016.xls` — on paths whose
instruction delta is +0.00% and −0.54%; two independent 150-sample repeats and
the final window put them at +2.78%/+0.62%/+3.12% and −4.36%/−1.70%/+1.02%
against floors reaching 11.43%, so they are reported as host spread rather than
an effect. One instrument disagrees with the others and is
reported anyway: callgrind scores `ConditionalFormattingSamples.xls` validation
at **+26.54%** instructions where `perf stat` scores it −45.85% and the clock
−41.89%, because callgrind counts `rep movsb` per byte (change 0604's 35×
overstatement) and the window copies that fixture's 1,328,049-byte stream once
more. No claim-registry entry; `performance_claim: none`.
[Change and limitations](0636-cfb-cursor-bounded-window.md);
[evidence](results/change-0636/README.md).

## For `ADR_COMPLIANCE.md`

## 0636 — a window that moves when bytes are read, not which bytes exist

The window is a **wrapper**, not a mode of `SharedOleStreamCursor`, and that is
the compliance argument as much as the performance one: every byte it moves goes
through the public cursor API, so the cursor keeps its type, its size, its drop
glue and its read path, and a caller that does not build a wrapper cannot pay
for one. ADR 0003's bounded-resource rule is met by a ceiling clamped to a
published 64 KiB constant, a window grown with `try_reserve_exact` that reports
`OleError::Allocation` on failure, and a fill clamped to the declared stream
length, so no byte outside the selected stream is ever read. ADR 0005's
lazy-payload contract is observed rather than altered: the window is filled on
demand from the same chain walk and never materializes a stream — the MiniFAT
parity test asserts the Mini Stream cache stays unmaterialized. ADR 0006 is
untouched: no execution context, no worker pool, no ambient I/O, no lock.
Because every fill is a cursor `read_exact`, change 0558's trailing fence and
change 0317's `SourceChanged`-wins precedence execute unmodified, and nothing is
committed before the step that can fail has succeeded. One contract does move,
for the opted-in caller only, and it is stated rather than absorbed: a read the
window covers takes no observation of its own, so change 0621's one-per-read
rule becomes one per fill and a mutation between two served reads is reported by
the next fill. Two things bound it — `validate_source_with_limits` ends every
exit path with an observation that refuses `SourceChanged`, so the
operation-level bracket is exactly what it was; and a mutation reverted before
the next fill is not observed, which `litchi_core::FileVersionPolicy` already
documents for reverted transitions. Change 0621's change-under-read sweep is
extended to fill boundaries and pins both halves. One defence forced the design
and is preserved to the byte: XLS validation must stop at a FILEPASS record
without reading the ciphertext behind it, so the window is not opened until
`filepass_slot_open` has closed and a properly-placed FILEPASS can no longer
appear — `xor-encryption-abc.xls` reads the same 8 requests and 2,078 bytes on
both legs. No ADR is amended and no ADR clarification is proposed.
[Change 0636](0636-cfb-cursor-bounded-window.md); `performance_claim: none`.
