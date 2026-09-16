# Log sections for change 0642

Four paragraphs for the coordinator to merge, one per log, in the style of each
log's newest sections. Nothing here edits those files.

## For `HOTSPOTS.md`

## 0642 — the XLSX range visitor stops building the range; XLSX-7 closed

Record: [0642](0642-xlsx-visit-cells-streaming.md).

**XLSX-7 is closed, and the `visit_cells` line in the 0587 XLSX section is now
wrong in the right direction.** The survey said the source-backed visitor "is
documented as a visitor but materializes the whole range first", modelling the
vector at over 5.2 MB for 65,536 cells. It did: `visit_cells` called `cells`,
took the `Vec<SourceCell>` (80 bytes each), and iterated it — and on the
materialized-store route it also cloned every `Cell` out of the store to build
that vector, only to hand the caller a reference to the copy. The route decision
and every source read now live in a private `select_cells` returning a private
`Selection`, which `cells` converts into exactly the vector it returned before
and `visit_cells` walks one cell at a time. **A whole-sheet walk over a
materialized store is now allocation-free**: 65,551 allocation calls,
5,608,448 allocated bytes and 5,608,448 peak live bytes fall to **zero** on a
2×256×256 integer probe workbook, and 16 calls with 10,485,760 bytes fall to
**zero** on `no_drawing_patriarch.xlsx`'s 75,770-cell worksheet. Instructions for
that walk fall **81.66%** (622.5 → 114.2 per cell) and **46.93%** (222.0 → 117.8
per cell); paired timing puts it at **+83.5% to +83.9%** and **+75.3% to +78.3%**
of p50 in both directions across two windows, against A/A floors of at most 3.9%
— 6.1× and 4.1–4.6× faster. The two corpora remove different things, which is the
useful part for ranking what remains: the dense probe's numeric cells cost one
`Box<str>` clone each (65,536 of its 65,551 calls), while the POI worksheet's
shared-string cells clone an `Arc` for free and all sixteen calls were the
vector's own geometric growth to a 131,072-slot capacity — a **42%
over-allocation that `cells` still pays**, because `collect_stored_cells` keeps
the original `try_reserve(1)` loop. Two things this does **not** move. A cold
whole-sheet read loses 5,242,880 bytes (exactly 80 × 65,536) on the scan route
and 10,485,760 on the stored route but its **peak live bytes do not change at
all**, because the peak is set by the worksheet payload plus the scan's own
record vector, or by the eager parse, before the removed vector is allocated; and
that cold read costs **+0.19% to +0.20% more instructions**, 29% of it the new
record-validation pass. The next item in this area is therefore the scan's own
`Vec<SelectedRecord>`, which is what still sets the cold peak — and it is a
contract change, not a refactor: making `scan_range` yield would move refusals
after the first visit, so it needs a frozen design record first. OLE2/OOXML
remain the active priority; ODF stays deferred and iWork excluded.

## For `GOAL_AUDIT.md`

## 0642 — a documented visitor made to behave like one, with the refusal order proved rather than assumed

Record: [0642](0642-xlsx-visit-cells-streaming.md).

`docs/GOAL.md` puts correctness and bounded resources above speed and says a
typed refusal may never be traded for a partial result, and this change is a
small test of exactly that rule, because the obvious version of it breaks the
rule. `SourceBackedWorksheet::visit_cells` built a whole-range
`Vec<SourceCell>` and then iterated it; the cheapest way to stop doing that is
to let the bounded scan yield cells to the callback as it parses. That would
move every refusal in the rows *after* the requested rectangle to after the
first visit, and the scan's own contract — "the returned eligible value is
published only after the shared MCE/XML stream reaches EOF" — is what makes
today's behaviour refuse-before-visit. **That version was not implemented.** What
landed keeps the scan running to EOF, keeps the dependency resolution and both
source/execution fences where they were, and *adds* a pass that validates every
retained record before the first callback, so the guarantee holds structurally
rather than by an argument about which arms are reachable. The two refusals that
pass hoists are provably unreachable — `SelectedCells` is only built by
`Scanner::finish` from `retain_selected`, whose two call sites each set exactly
one of the two fields — and they were hoisted anyway, so that a future scanner
change cannot quietly start refusing mid-walk. The audit should note the price:
that pass costs 13 instructions per record, 852,085 per 65,536-cell cold read,
and the cold path is **+0.19% to +0.20%** of instructions overall. It should also
note what the change buys on the resource axis the goal document cares about: the
warm walk's peak live bytes go to zero, which is a bounded-resource improvement
rather than a latency one, while the **cold** read's peak does not move at all
because the scan's own record vector still sets it. Evidence tiers: **measured**
for every allocation, byte, peak, instruction count and timing quartile, and for
the 397-worksheet differential; **modelled**: nothing; **unknown**: whether the
scan's record vector can be removed without moving a refusal, which is the
frozen-design question this change deliberately left open. One gate reports
warnings and they are pre-existing: `cargo clippy -p litchi --features
docx,xlsx,pptx,xls --all-targets` emits six warnings (`passing a unit value`,
one unused function, one needless `mut`) reproduced with identical text on the
untouched before checkout. `performance_claim: none`. OLE2/OOXML remain the
active priority; ODF stays deferred; iWork is excluded.

## For `REPORT.md`

## 0642 — a whole-sheet visit that allocates nothing, and the peak it does not move

Two operations on the XLSX source-backed door read a rectangle of cells:
`cells`, which returns an owning `Vec<SourceCell>`, and `visit_cells`, which
hands each cell to a callback. Until this change the second was the first plus a
loop — it built the whole vector, then iterated it — so a visitor that was
documented as never outliving the read paid 80 bytes per cell for a copy it
threw away, and on the materialized-store route it cloned every cell to make
that copy. This change gives the two operations a shared private route selector
and lets the visitor produce one cell at a time. On a walk over a worksheet whose
store is already materialized — the thing a visitor is for — the operation now
allocates **nothing at all**: zero allocation calls, zero bytes, zero peak live
bytes, down from 65,551 calls and 5,608,448 bytes on a 65,536-cell integer sheet
and from 10,485,760 bytes on `no_drawing_patriarch.xlsx`'s 75,770-cell
worksheet. That is 81.7% and 46.9% fewer instructions and a 6.1× and 4.1–4.6× p50
speedup, measured in both directions of an A1 B1 B2 A2 pair in two windows, with
an A/A floor of at most 3.9% in each. The report should be equally clear about
the two boundaries. First, on a **cold** read — open, scan, resolve, convert —
the same 5.2 MB or 10.5 MB disappears from the allocated-byte total but the
**peak live bytes do not change**, because the scan's own record vector and the
worksheet payload set the peak before the removed vector exists; a reader who
wants the cold peak down is waiting for a different change. Second, that cold
read is **0.19% to 0.20% more instructions**, and the record reports it rather
than averaging it away: 29% is the new pass that validates every retained record
before the first callback, which is what keeps a refusal in a later row from
arriving after a cell has been visited. Value identity is not asserted but
measured: a four-way differential — cold `visit_cells`, cold `cells`,
warm-store `visit_cells`, and the eager `Workbook` store — over **397 worksheets
of 182 `.xlsx` files** produces one table per binary, and the two tables are
byte-identical, including the thirteen worksheets that refuse and the exact text
of each refusal. `performance_claim: none`; the numbers above are evidence.
OLE2 and OOXML remain the active priority; ODF is deferred; iWork is excluded.

## For `ADR_COMPLIANCE.md`

## 0642 — no boundary moved, and the one that could have was left alone

ADR 0003 makes borrowed views the fine-traversal idiom and requires conversion
to owned storage to be explicit; ADR 0005 makes semantic payloads lazy behind
caches whose behaviour is semantically invisible; ADR 0006 makes `Preserve` the
default and forbids trading a refusal for a partial result. This change is
compliant on all three, and the interesting part is the third. `visit_cells`'s
documented promise is that "no callback runs while a source reader is active",
and it kept that promise the blunt way, by reading the whole range into an owned
vector first. It now keeps it precisely: the private `Selection` it walks holds
either a borrowed reference to the already-published worksheet store or records
the bounded scan produced after `with_verified_decoded_reader` returned and the
dependency readers were released, so neither variant can be alive while a reader
is. The `finish_result` fence that published the old vector now publishes the
selection, in the same place, so a source mutation or a cancellation still
outranks a callback's own error; cancellation is still checked before every
callback. **The ADR 0006 boundary that could have moved did not.** The bounded
scan still reaches worksheet EOF before anything is published, so a malformed
row *after* the requested rectangle still refuses before the first visit — and a
new pass validates every retained record before that visit too, so the guarantee
is structural rather than a reachability argument. Nothing was relaxed: the same
`StreamLimits`, the same `Capabilities`, the same `selected_stream_limits`
ceiling, the same four dependency-fallback conditions, the same
`validate_styles`, the same `invalid(...)` texts in the same order. On ADR 0003,
the visitor now hands the store's own `&Cell` to the callback on the stored
route, which is the borrowed-view rule applied where it was previously
side-stepped, while `cells` remains the explicit owning conversion and returns
the identical vector. No public API, error type, limit, defence or output byte
changed; no new `unsafe`; no new dependency; no ambient I/O; no executor. The
compliance statement to carry forward is the one the change declined to make:
letting `scan_range` yield to the callback would relocate a typed refusal past a
partial result and is therefore an ADR 0006 change, not an optimization, and it
needs a frozen design record before anyone implements it.
