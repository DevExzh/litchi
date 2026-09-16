# Log sections for change 0647

These four blocks are written for the coordinator to merge into the shared
program logs. Each is in the style of that file's newest section.

## For `HOTSPOTS.md`

## 0647 — the one `litchi-opc` finding 0628 left open, closed: a relationship call that establishes nothing keeps its pristine proof

Record: [0647](0647-opc-get-or-add-noop-reuse-design.md).

**0628's parting finding is fixed, and the contract question it stopped for
turns out to have been answered already by the code that fix
replaces.** `Relationships::get_or_add` reuses an existing relationship by
handing the established identifier to `add_relationship`, whose first statement
was `self.invalidate_source_capture()` — so a call that establishes nothing
destroyed change 0593's open-time proof, and a `.rels` member 0593 would have
copied verbatim was reserialized, audited and byte-compared instead. 0628
declined to fix it because keeping the proof looked byte-visible: wherever the
source spelling of a `.rels` member differs from this crate's canonical
serialization, the published bytes would move from canonical to source. **The
corpus says that would be almost everywhere and the code says it happens
nowhere.** A probe over all 336 OOXML fixtures finds **2,189 of 2,216
relationships members spelled differently from canonical (98.78%)**, carrying
6,214 of the corpus's 6,281 internal relationships; and a second probe performs
all **6,281 reuse calls**, publishes each one and its seam-only baseline through
the owned-source route, and finds **6,281 of 6,281 byte-identical** with
**0 differing**, on both legs, with an **empty cross-leg `diff`** over per-fixture
rolling SHA-256 digests. The reason is in `try_write_preserved`: the
serialize-and-compare branch compares the canonical bytes it just built against
the *same* open-time capture and, on equality, takes
`PreservationAction::Copy` of the **source** archive entry — precisely what the
pristine branch takes. The fix is one `match`: `add_relationship` destructures
`self`, and only the vacant arm inserts and only the vacant arm clears the
capture, which also keeps the argument strings out of the occupied path. Per
save of the 132-member `ConditionalFormattingSamples.xlsx` with one reusing
call, `try_to_xml_bytes` falls **42 → 41**, `verify_authored` **1 → 0**, a part
owner's `rels_uri` **225 → 224**, save allocations **402 → 363** (package owner)
and **523 → 363** (`drawing1.xml` owner), and the canonical `.rels` bytes built
and discarded fall **733 → 0** and **2,260 → 0**. Instructions per open + reuse +
save fall **−0.078%** and **−0.450%** against seam-only controls that move
−0.030% and −0.008% on recompile alone; the *marginal* cost of the reusing call
over a bare seam falls **+23,176 → +10,294** and **+109,387 → −7,363**
instructions. Paired p50 is −0.69% (package) and −1.40% (drawing, second window)
against A/A floors under 0.09%, but the removed work is only 0.08–0.45% of the
scenario, so no timing claim is made; the drawing scenario's first window is
reported with a **+11.70% mean and +230.79% p99 regression** traced to one
contaminated block at load average 39 and chased by the re-run. **What this does
*not* buy today** is anything on a format's ordinary save: the DOCX route,
measured over 63 fixtures, publishes byte-identically on both legs, because its
save rebuilds `word/document.xml` into a fresh `BlobPart` that never carried a
capture, and 23 members across 52 published packages — 21 of them
`word/_rels/document.xml.rels` — are regenerated for that reason both before and
after. The gain is at the `litchi-opc` seam —
`Part::relate_to` and `OpcPackage::relate_to` are public — and, per an audit of
every production call site, at two `litchi-xlsb` "ensure" helpers that guard on
the part rather than the relationship. `performance_claim: none`. OLE2/OOXML
remain active; ODF is deferred until completion and iWork excluded.
[Retained evidence](results/change-0647/README.md).

## For `GOAL_AUDIT.md`

## 0647 — preservation-by-default reaches a case it was already reaching, and the rule-1 ordering holds

`docs/GOAL.md` puts correctness and lossless preservation above speed and orders
optimization as *eliminate unnecessary work first*. Change 0647 is that first
rung with nothing traded for it: the work removed is a canonical `.rels`
serialization and its publication audit that were built, compared against an
identical copy of themselves, and thrown away. The audit question the change
raises is whether preservation moved, and the answer measured over 336 OOXML
fixtures is that it did not and could not: 6,281 reuse publications are
byte-identical to their no-op baselines on both legs, and all 96 baseline and 260
reuse refusals are the same two typed kinds (`PreservationUnavailable`,
`SignedSourceRequiresExplicitPolicy`) in the same counts. That matters for the
goal because the premise under which change 0628 declined the fix — that keeping
the proof would publish the source spelling in place of the canonical one —
would have been a *preservation improvement* that nonetheless changed bytes, and
the rule is that a byte change is a contract change whichever direction it moves
in. The design was therefore written frozen and only implemented once the
measurement proved the question moot; the record states the rule, the byte
consequence, the admission gates and the count that satisfied them, so a future
reader can see the design that would have been frozen had the count come out the
other way. Two residual gaps are recorded rather than closed: a pristine member
no longer runs `try_to_xml_bytes` or `verify_authored` on the preservation route,
which is change 0593's already-recorded behaviour class and needs roughly 62,500
relationships in one member against a **measured corpus maximum of 43**; and the
ordinary format routes gain nothing today, measured for DOCX and audited for
XLSX, PPTX, XLSB and `litchi-ooxml-common`. Every crate depending on
`litchi-opc` was tested, plus the feature-bearing `litchi` suite and the
harness's own. `performance_claim: none`. OLE2 and OOXML remain the active
priority; ODF stays deferred; iWork is excluded.

## For `REPORT.md`

## 0647 — the frozen design that measured its way into an implementation

Change 0628 ended with one unfixed `litchi-opc` finding and an explicit reason:
making `Relationships::get_or_add` keep change 0593's open-time capture on a
reuse would turn a reserialized `.rels` member back into a copied one, and
wherever the source spelling differs from the canonical serialization that is a
published-byte change. The brief for 0647 asked for a frozen design record and an
implementation only if the corpus showed zero fixtures whose bytes would move.
The corpus shows that, but not for the reason anyone expected. **The byte
consequence does not exist at all**, because the route the capture short-circuits
does not publish the canonical bytes it builds: `try_write_preserved` compares
them against the same open-time canonical capture and, when they are equal,
takes `PreservationAction::Copy` of the source archive entry — the source
spelling, its compression method and its local framing, verbatim. The pristine
route takes the identical action. So a non-canonically spelled `.rels` already
republishes as the source's own bytes today, and the capture only decides whether
the canonical form is built and audited before that conclusion is reached. The
full-writer route is symmetric: `materialize_pristine` serializes and audits
every pristine member before `write` emits a byte, so that route emits the
canonical form on both legs. The implementation is one `match` in
`add_relationship`, not in `get_or_add`: the entry now decides both the value and
the proof, so the collection changes if and only if the capture is dropped, and
the reuse path stops building and discarding a `Relationship`. The evidence is
three corpus sweeps with empty cross-leg `diff`s — 6,281 reuse publications over
334 fixtures, 2,216 spelling comparisons, and the DOCX ordinary route over 63
fixtures — plus four in-crate integration tests that pin the equality on a fixture
whose members are *provably* non-canonically spelled, so the tests cannot
silently stop testing the thing they were written for. The record states what is
not bought: nothing on a format's ordinary save today, and two refusals skipped
on a path no fixture reaches. `performance_claim: none`.

## For `ADR_COMPLIANCE.md`

## 0647 — no boundary moved; the preservation proof is narrowed to exactly the operations that can invalidate it

**ADR 0006** requires preservation by default — untouched entries, ordering,
compression, timestamps and lexical details retained when possible — and
deterministic serialization absent an explicit `Clock`, actor identity or
cryptographic RNG. Change 0647 removes a generate-and-discard step and publishes
the same bytes; the published output remains a function of the package, proved
over 6,281 reuse publications across 334 fixtures with an empty cross-leg `diff`
and identical SHA-256s in the counts probe. **ADR 0005's** 2026-08-21 amendment
makes preservation provenance planning evidence only, never authorization for
exact passthrough: the capture is used here exactly as change 0593 uses it, to
choose `Copy` for one member inside an already-proven preservation plan, and
`exact_source_authorized` is untouched — every seam that reaches the changed
method (`OpcPackage::get_part_mut`, `OpcPackage::rels_mut`,
`OpcPackage::relate_to`) still revokes it first, so nothing widens who may take
the whole-archive passthrough. **ADR 0003** is untouched: no snapshot, patch or
source-preservation proof changes. The compliance statement to record is about
the *proof's* scope rather than any boundary: the invariant was "every mutating
method clears the handle", and it is now stated precisely — every method that
*changes the collection* clears it, and a call that establishes nothing does not,
because the handle's contract is that it describes the collection's value. The
conservative direction is preserved: the handle is cleared on every path that
could have changed the value, and a cleared handle costs only the
serialize-and-compare route. Three fallible steps sit behind the proof;
`PackURI::rels_uri` → `InvalidPackUri` does **not** move (the preservation route
still derives it in its final-member-name loop and the full writer in
`materialize_pristine`), while `try_to_xml_bytes` allocation failure and
`verify_authored` on a discarded serialization are skipped, which is change
0593's already-recorded intentional behaviour class, reachable only above roughly
62,500 relationships in one member against a measured corpus maximum of 43.
No new `unsafe` (the crate is `#![forbid(unsafe_code)]`), no limit relocated or
weakened, no malformed-input defence removed, no ambient I/O, no global state or
executor, no public type added, and no archive type, raw lock or executor leaked.
The one widening is recorded: the rule applies to any `add_relationship` call
naming an already-taken identifier, which was already a silent no-op because the
occupied arm never replaced anything.
