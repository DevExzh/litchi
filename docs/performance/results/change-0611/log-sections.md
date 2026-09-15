# Log paragraphs for change 0611

Four paragraphs, one each for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and
`ADR_COMPLIANCE.md`, in the style of their newest sections. The coordinator
merges these; this change does not edit those files.

## HOTSPOTS.md

**ZIP-2 is closed.** A member's first read on a positional source cost two or
three requests — the 30-byte fixed local header, the payload, and a data
descriptor when one is declared — and the framing between the header and the
payload was skipped rather than read. Change 0611 takes one bounded positional
read of the member's whole local record on the ordinary indexed read path and
answers every read of that member from it. Measured on change 0587's own probe:
a source-backed open of the 132-member workbook falls from 87 requests to 45,
`shapes.pptx` from 45 to 24, `comment.docx` from 12 to 6, a first part read from
2-3 to 1, and a whole 90-part traversal from 208 to 90, for +0.5% to +23% bytes
— the local variable region, 84 bytes per structural member on the workbook.
0587 modelled 87 → 45, 45 → 24, 12 → 6 and 2-3 → 1; all four are met exactly.
On the 1 ms-per-request transport of changes 0493 and 0572 the request drop is
wall clock. Rank 14 of the 0587 queue is done; **ZIP-5** (the accessor change
0577's coalesced structural prefetch is blocked on) is now the next ZIP item and
its modelled 45 → about 10 is measured from this base, not from 87. ZIP-1 and
ZIP-3/4 were taken by changes 0594 and 0600.

## GOAL_AUDIT.md

Change 0611 is the first change in this programme to price a read-grammar change
on a latency-bearing transport rather than on a warm local source, and it is
worth recording why that was the right instrument. Change 0561 attributed this
span model and declined to implement it because change 0493 had already measured
the same mechanism, as an opt-in read-ahead window, at ±2% on a warm local
source and −85% at 1 ms per request; its conclusion was that collapsing these
reads is worth nothing locally and must be argued against a range-source
measurement. This change is the measurement 0561 asked for, and it confirms both
halves: on local, in-process sources every selector sits inside the A/A floor and
nothing is claimed, while on the simulated transport a source-backed OPC open
halves its wall clock. The standing instruction that instruction counts rank
work and not latency has a counterpart here — **request counts rank latency on a
transport and not work**; the deterministic count is the durable result and the
timing is its price at one service time. A third lesson is about *test triggers* rather than assertions: sixteen tests in
two consumer crates identified "the read that fetches this member" by its start
offset, and every one of them broke without a single contract moving. Three
successive predicates were needed before one held under both grammars, and each
wrong one was found by a different test — a zero-length terminating read landing
on the next member's header, a preservation probe at the same offset as the
member read, and a bulk publication copy spanning the member from 64 KiB away.
A read-grammar change should expect to spend as much effort on triggers as on
the code.

The second lesson is about oracles. The
differential change 0582 built, the gate 0587 names for any read-grammar change,
exercises only the strict-layout APIs and never calls `IndexedArchive::read_entry`
— the path every ordinary OPC part read takes. Re-running it unchanged would
have proved nothing about this change. It had to be extended before it could
gate, and once extended it immediately found a real divergence that no test in
either crate caught.

## REPORT.md

**0611 — one bounded positional read per ZIP member first-read.** Retained,
implemented in `soapberry-zip` (`archive.rs`, `office.rs`), `performance_claim:
none`. Deterministic counts on change 0587's probe: open 87 → 45 requests on the
132-member workbook, 45 → 24 on `shapes.pptx`, 12 → 6 on `comment.docx`, a first
part read 2-3 → 1, a 90-part traversal 208 → 90, for +0.5% to +23% bytes and
+1.3% to +4.5% allocation calls; source observations, the `stream_to` counts and
forward-versus-reverse behaviour are unchanged. Paired ABBA timing on the 1 ms
per-request simulated transport of changes 0493 and 0572, 30 samples per leg,
CPU 31, with an A/A floor in the same window under 0.4% at p50. Correctness: the
change 0582 differential, extended with `IndexedArchive::read_entry` and a
slice-backed control and re-run in full — both builds, 22,875 archives, every
member, both directions, both limit profiles, 2,886,786 member verdicts. Its
first run found two class-E divergences on one crafted mutation whose central and
local records disagree about compressed size: the same member was refused on both
legs with a different typed error, because a partly-covered read re-chunked the
decoder's input and `ZipVerifier` completes a member on a per-read size test. The
design was amended so the buffer serves a payload only when it holds all of it,
and the re-run is clean. Thirteen new tests in `soapberry-zip`; sixteen
tests in `litchi-opc` and `litchi-xlsx` whose *trigger* encoded the old read
grammar were corrected with every assertion unchanged. Gates: ten sections, all
exit 0, including the consumer suites for XLSX, DOCX, PPTX, XLSB, ODT, ODC and
ODF.

## ADR_COMPLIANCE.md

**ADR 0005 (bounded resources, read-order independence).** The speculative read
is bounded by a named ceiling, `MAX_MEMBER_SPAN_READ_BYTES` = 64 KiB, measured
against 7,757 OOXML members and equal to the largest window `litchi-opc` already
admits; the buffer is reserved with `try_reserve_exact` and lives only for the
member read, so the resident bound added is one buffer per member read in
flight. A member above the ceiling takes no speculative read at all. Read-order
independence is measured, not asserted: forward and reverse traversal of all 90
parts of the 132-member workbook cost identical requests and bytes on both legs,
and the differential runs every member under both directions. **ADR 0006
(lossless preservation).** No output byte is produced on this path and no
preservation surface is touched; the preservation index and the writer are
unchanged. **ADR 0011 (ownership).** No archive type crosses a crate boundary:
the window is computed inside `soapberry-zip` from central facts and physical
order the index already holds, which is why `litchi-opc` needed no production
change — in contrast with change 0577's candidate (c), which still needs a new
read-side accessor. **Change 0317 (error precedence).** The change sits entirely
inside 0317's brackets; source-version failure still precedes execution-context
failure, which still precedes the mapped ZIP member error, and the span read is
tried once so one member read stays one refusal, one cancellation observation and
one resource reservation. **One stated consequence.** Under an `ExecutionContext`
with a finite `Resource::InputBytes` limit, a spanned first read commits the
member's local variable region as well — bounded by 610 bytes per member,
measured at 84 — so a finite input budget can be exhausted marginally sooner. No
error identity changes and no new refusal exists. This is the same class and
direction as the speculative window change 0573 already landed on the strict
path.
