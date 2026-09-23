# Log sections for change 0753

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself; the numbers
are the record's.

---

## For `HOTSPOTS.md`

## 0753 — The legacy fresh writers encode each text once into its final buffer and hash each shared string once

[0753](0753-legacy-fresh-writer-text-paths.md) removes the three fresh-writer
costs profile r2 priced. The DOC writer counted, walked and appended every
character separately (three passes, 60% of `doc_fresh_write_to`): it now counts
UTF-16 units without decoding, widens ASCII through a `TrustedLen` iterator,
walks characters for field CPs only when a vectorized byte scan finds one, and
writes the text once, straight after the FIB placeholder in a stream reserved
from the model, instead of growing a separate buffer and copying it twice. The
PPT writer re-counted ASCII text as UTF-16 and copied each text byte about
fifteen times through record builders; its text box, shape, drawing and slide
records are now written in place into the document stream. The XLS writer
hashed each shared string three times plus once per map resize and cloned it
twice; a per-write table now borrows the strings, caches each keyed SipHash-1-3
once and records every cell's index. Output is byte-identical (base-captured
goldens, a differential SST test and an in-process order test). Paired ABBA on
CPU 8, payload-heavy: DOC 0.286, PPT 0.191, XLS 0.374 of base time; large: DOC
0.787, PPT 0.527; instructions per write fall to 0.091, 0.160 and 0.425.
Remaining: DOC is now mostly the CFB writer's destination `memset` (37%) and
first-touch page faults of output-sized buffers; PPT clones each text into
`UserShapeData` (15% with the conversion); XLS pays one keyed SipHash per string
cell (27%). Found, not fixed: the fresh XLS SST order follows a per-process
`HashMap` seed for sheets with several distinct strings.

---

## For `REPORT.md`

## 0753 — Single-pass text and hash-once shared strings in the legacy fresh writers

[0753](0753-legacy-fresh-writer-text-paths.md) retains six production commits
in `litchi-doc`, `litchi-ppt` and `litchi-xls` (writers only; no public API,
format, limit, dependency or `unsafe` change; litchi-cfb untouched) and one
test commit. Paired ABBA, two windows, CPU 8, final build `4c079ef4bf` against an
identically built base: `doc_fresh_write_to` payload-heavy 5.683 → 1.628 ms
(0.286), large 0.787, tiny 0.883; `ppt_fresh_write_to` payload-heavy 4.504 →
0.860 ms (0.191), large 0.527, tiny 0.910; `xls_fresh_write_to` payload-heavy
6.227 → 2.321 ms (0.374), all-numeric large 0.915/0.951 and tiny 1.011/1.014
(+0.7% and −0.1% exact instructions). Callgrind instructions per write: DOC
0.091, PPT 0.160, XLS 0.425 on payload-heavy. Allocated bytes fall 13–73% on
every DOC and PPT shape and 30% on XLS payload-heavy; no peak live heap rises
(XLS payload-heavy −38%). Controls stay within ±1.1% except
`doc_semantic_one_edit_save` large, 0.827 with identical instructions (heap
state left by the changed corpus writer; not claimed). All outputs are
byte-identical: 24 base-captured golden fixtures (ASCII, Latin-1, CJK,
supplementary-plane, empty and >64 KiB text, fields beside surrogate pairs,
every DOC story, PPT interactions, tables, pictures, notes through `save`, XLS
strings across CONTINUE records and past 0xFFFF units) pass on both legs.
`performance_claim: none`.

---

## For `GOAL_AUDIT.md`

## 0753 — redundant text work removed from the fresh writers with every byte, refusal and hash guarantee kept

[0753](0753-legacy-fresh-writer-text-paths.md) keeps the fresh writers' output
byte-identical and deterministic where it was, keeps every limit and typed
refusal in its order (length checks still precede the work they guard; the
overflow checks of the skipped per-character walks are made once with the same
messages; a PPT write refused after a slide's drawing is already in the stream
leaves the destination untouched), keeps infallible growth where it was, and
keeps hash-flooding resistance: the XLS table caches the same randomly keyed
SipHash-1-3 `std`'s map computes and passes it through, with a collision test.
No ADR, owner decision, public API, format or dependency changed. It records a
pre-existing ADR 0006 gap for follow-up: the fresh XLS writer's SST order for a
worksheet with several distinct strings depends on a per-process `HashMap`
seed. The non-iWork goal remains active.
