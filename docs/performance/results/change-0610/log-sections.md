# Log sections for change 0610

Four paragraphs for the coordinator to merge, one per document, in the style of
each document's newest section. Each is written to stand alone. Their links are
relative to `docs/performance/`, where the four log documents live, not to this
packet directory.

## For `docs/performance/HOTSPOTS.md`

## 0610 — an XLSX open-edit-save reads one part of ninety

Design only, implementing 0587's SAVE-5 and drafting the proposed ADR that gate
1 of 0581 requires. 0581 measured what the eager `OpcPackage` *retains*; nobody
had measured what an operation *reads*. A retained ablation probe now does:
substitute a sentinel payload for one part at a time and compare the operation's
outcome everywhere except that part's own member. `load_parts_eager` inflates
**5,077 parts and 45,562,463 bytes** across 334 real OOXML fixtures whose
archives total 15,323,729 (2.97×; XML parts are 90.8% of the inflated bytes).
The documented XLSX editor route — `Workbook::open` → `edit()` → hide a tab →
`commit()` → `to_bytes()` — reads **exactly one part on all 33 corpus fixtures
that admit it, `/xl/workbook.xml` every time**: 33 of 550 parts and 60,239 of
8,274,037 inflated bytes, **0.728%**; on the 132-member
`ConditionalFormattingSamples.xlsx` it is 1 of 90 parts and 0.164% of bytes. The
physical `OpcPackage` open-edit-publish route reads **zero** other parts across
325 fixtures and 4,896 parts, so the publication plan needs no decoded payload
for any member it copies — which is true only because 0593 replaced the plan's
byte comparison with a provenance-identity proof, making 0593 a prerequisite for
C2′ rather than merely a predecessor. The census also measures Σ decompressed
payloads directly and confirms 0581's stated inference: `archive + Σ` accounts
for **98.32%** of the retention 0581 measured across its six axis-3 fixtures.
The migration surface is far smaller than 0587 estimated: of 307
`iter_parts()` sites under `crates/*/src`, 21 are inline tests and 27 are
`SourceBackedPackage::iter_parts` (out of scope), leaving **259 in-scope
production sites**, of which the payload-reading group is bracketed **16–88** by
two mechanical scopes and put at **about 25** by an independent full-context
review; none of the 1,078 `.blob()` sites changes, because `Part::blob` keeps
its signature. Gate 2 is also smaller
than 0581 assumed — the reader already charges `PartBytes` and `TotalPartBytes`
against declared central-directory sizes before any decompression and again
against actual sizes after, so only the second charge moves.
`performance_claim: none`; `claim_authorized: false`; nothing is implemented and
ADR 0030 is **proposed, not accepted**. Also reported, not diagnosed: the
pre-existing `NotCompact` defect 0587 §4 found on two operations of one fixture
blocks the `hide` operation on **59 of 180** `.xlsx` fixtures. Remaining in this
area: SAVE-3 (PPTX eager save regenerates every slide), SAVE-4 (fresh deflate
state per regenerated member), and C3 as the end state. OLE2 and OOXML remain
active; ODF is deferred until that goal completes and iWork is excluded.
[Change and limitations](0610-opc-lazy-part-decode-design.md);
[proposed ADR](../adr/0030-lazy-opc-part-decode.md);
[retained evidence](results/change-0610/README.md).

## For `docs/performance/GOAL_AUDIT.md`

## 0610 — hypothesis 8 is confirmed, and the waste is measured

Design only, against the audit's standing eager-OPC-materialization row.
`docs/GOAL.md:350`'s hypothesis 8 — *"XLSX selective sheet loading may be
undermined by eager OPC materialization"* — is now **confirmed by measurement
rather than by source reading**: the XLSX layer above OPC already defers
worksheet work, and an open-then-hide-a-tab-then-save reads exactly one part,
`/xl/workbook.xml`, on every one of the 33 corpus fixtures that admit the
operation, while `load_parts_eager` has already inflated all 550 parts and
8,274,037 bytes before the workbook model exists. 99.27% of those bytes are
decompressed, charged, retained and never read. The record also closes a gap
0581 left explicitly open: 0581 inferred that its measured open retention is
`archive + Σ decompressed payloads` and said the inference was not a
measurement; the new census measures Σ and the formula accounts for 98.32% of
0581's figure. `GOAL.md`'s standing instruction — *"Do not change an
architecture solely because it appears suboptimal in source code. Require
profiles and scenario measurements"* — is honoured: the measurements exist and
the architecture is still not changed, because gate 1 requires a human to accept
[ADR 0030](../adr/0030-lazy-opc-part-decode.md) first. Audit rows this does
**not** close: no DOCX or PPTX semantic-editor read set is measured, because no
example opens a real `.docx` or `.pptx`, edits through the model and saves —
0587 §4 recorded that gap, 0593 recorded it again, and it remains open; a DOCX
save regenerates every model-owned part when one whole-model flag is set and a
PPTX save regenerates every slide, so those read sets are **unknown, not small**.
`tools/perf-baseline` still has no ordinary open-and-save selector. One new
audit item is opened: the pre-existing `NotCompact` publication refusal blocks a
plain tab-hide on 59 of 180 real `.xlsx` fixtures, which is a correctness defect
sized here and owned by no record. `performance_claim: none`, and no timing,
peak-RSS or cold-cache figure is measured. OLE2 and OOXML remain active; ODF is
deferred until that goal completes and iWork is excluded.
[Change](0610-opc-lazy-part-decode-design.md);
[evidence](results/change-0610/README.md).

## For `docs/performance/REPORT.md`

## 0610 — a proposed ADR for lazy OPC part decode, with the design frozen and sized

No file under `crates/` changed. Three documents are added: the **proposed, not
accepted** [ADR 0030](../adr/0030-lazy-opc-part-decode.md), deliberately absent
from the accepted ADR table and listed instead in a new "Proposed records (not
accepted, not normative)" section beneath it; a frozen design record; and an
evidence packet with the sizing probe. The proposed design keeps
`Part::blob(&self) -> &[u8]` **infallible**, so none of the 1,078 `.blob()` call
sites changes, and instead relies on the invariant that no `&dyn Part` is handed
outside `litchi-opc` before its payload is decoded — `get_part`, `get_part_mut`,
`main_document_part` and `part_by_reltype` already return `Result` and would
force the decode; `iter_parts` is the only infallible route and gains a
`try_iter_parts` sibling while its own item type narrows to a metadata view with
no `blob()`, so the invariant becomes a compile-time property. Payload state
lives in a `std::sync::OnceLock<Arc<Vec<u8>>>`, the only interior-mutability
primitive that keeps both `#[derive(Clone)]` and `Send + Sync` on `OpcPackage`;
`Cell`, `RefCell` and `OnceCell` cost `Sync` and `RwLock` has no `Clone` impl at
all. The record enumerates the migration surface (259 in-scope production
`iter_parts()` sites; the payload-reading group bracketed 16–88 by two
mechanical scopes and put at about 25 by an independent full-context review),
states the exact provenance interaction with 0593, and gives the predicted
retention on 0581's axes with one measured decoded-bytes point inside it. Two
refusal relaxations are put to the reviewer explicitly, both reachable only with
a central directory that under-declares a part, because the declared-size charge
already happens before decompression and does not move. Validation: nothing
changed, so no limit, audit, refusal, fence or signature moved.
`cargo fmt --all --check`, `check_report_claim_classification.py`,
`check_crate_boundaries.py` and a 51-link relative-link check all pass;
`check_perf_claims.py` and `check_example_targets.py` are red on the untouched
base and were reproduced byte-identically there. No latency, RSS, cold-cache or
throughput claim follows, and no implementation is authorized. See
[Change 0610](0610-opc-lazy-part-decode-design.md); `performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## 0610 — a proposed ADR, and gate 2 turns out to be smaller than it looked

Nothing is implemented, so no accepted record's position changes; what this
batch adds is a **proposed** ADR for review and the evidence a reviewer needs.
ADR 0005's "Input and lazy state" clause — *"Semantic payloads load lazily into
thread-safe weighted caches"* — actively favours the proposal, and 0581 already
recorded that no accepted ADR requires eager part decoding, an absence noted
there as an absence of documented rationale rather than as permission; this
record does not upgrade that absence into permission either. ADR 0005's
2026-08-21 exact-source amendment is untouched: the retained source archive
stays, `exact_source_authorized` is not widened, and the proposal uses
provenance exactly as 0593 already uses it — as planning evidence choosing
`Copy` for one member inside an already-proven preservation plan, now with
"never decoded" as the proof instead of a byte comparison. ADR 0005's
hierarchical-budget clause is where the one real question lies, and it is
smaller than 0581 framed it: the reader **already** charges
`ReadResource::PartBytes` and `TotalPartBytes` against declared
central-directory sizes before any decompression, with a regression test
asserting that the bulk decompressor is never invoked when that pass fails, and
charges them again against actual sizes afterwards. Only the second charge would
move, so every package refused at `open()` today for declared size is still
refused at `open()` with the identical `OpcError::ReadLimit` value and text. The
two resulting relaxations are stated for the reviewer rather than assumed: a
forged central directory's part is refused at first access instead of at open,
and the aggregate actual charge covers decoded parts rather than all parts. The
conservative alternative — holding the declared aggregate as a lifetime
reservation — is already implemented at the physical layer
(`reserve_declared_parts` / `commit_actual_parts` / `release_declared_parts`)
and is offered rather than decided. ADR 0003's panic-free facade requirement is
met by forcing the decode in the fallible accessor rather than by a new
`Result`; ADR 0006 and record 0528 are untouched, because no audit of changed
XML moves; ADR 0011's ownership boundary is untouched, because the seam is
entirely inside `litchi-opc`. No new `unsafe`, limit change, weakened defence,
global cache, executor, lock, Rayon pool, ambient I/O or leaked archive type is
proposed. `docs/adr/README.md` gains a "Proposed records (not accepted, not
normative)" section stating that nothing may cite ADR 0030 as authority until a
human accepts it. See [Change 0610](0610-opc-lazy-part-decode-design.md);
[ADR 0030](../adr/0030-lazy-opc-part-decode.md); `performance_claim: none`.
