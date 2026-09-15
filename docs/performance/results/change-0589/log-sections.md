# Log sections for change 0589

Four paragraphs for the coordinator to merge, one per log, in the style of each
log's newest sections. Nothing here edits `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` or `ADR_COMPLIANCE.md` directly.

---

## For `HOTSPOTS.md`

### 0589 — DOC and PPT source-backed opens: half the SHA-256 work removed

Survey item DOC-1 (rank 2 of change 0587) is priced and its value-identical
subset is implemented. `overlay::fingerprints` and `write_validated` drove two
SHA-256 hashers over one complete source read; with an empty span list
`apply_spans` cannot write a byte, so the target hasher consumed exactly the
bytes the source hasher consumed and finalized to the same digest. That covers
every DOC identity pass, both `ensure_current` passes, the whole PPT text-edit
snapshot open, and every exact byte no-op publication. Hash passes per open fall
from **12.00 to 6.00** for the generic DOC open and **4.00 to 2.00** for the PPT
text-edit open, counted exactly from callgrind's constant software-SHA cost of
52.07 Ir per artifact byte. Native `perf stat` isolation pairs on all 38 admitted
fixtures give cycles **median −36.6%** (range −48.0% to −25.5%) and instructions
median −29.0%, every fixture improving; paired timing gives p50 −48.1% on
`picture.doc`, −43.4% on `duplicate-style-names.doc`, −48.1% on
`cryptoapi-proc2356.ppt` and −47.9% on `45543.ppt`, against an A/A floor of
p50 ±0.1% and p99 ≤ 2.5% in the same window. Complete source reads are
**unchanged byte for byte**: the change removes hashing, not reading. The
remaining DOC-1 items are the open's third identity pass (predicted: 2 of the
6 surviving SHA-256 passes and 2 of the 6 complete reads) and the duplicate CFB
index parse (measured at 23,023–51,326 cycles, 6.7%–0.27% of the post-change
open); both are read-twice-compare defences whose removal would change when a
typed refusal happens, so both are designed and left in place. Change 0587
explained change 0586's zero by this hashing term; that reading needs a
correction, because 0586's zero was **structural** — the DOC paragraph hint
removed 0 of 3,991 chain links on 8 of 8 fixtures, so there was nothing for the
hashing to mask. What this change does is shrink the denominator: the hashing
share of a native DOC or PPT open falls from 51.0%–95.9% of cycles to
34.3%–92.2%, so a future attempt at those paths has roughly twice the
attributable headroom it had.
[Change 0589](0589-ole2-snapshot-fingerprint-passes.md);
[evidence](results/change-0589/README.md); `performance_claim: none`.

---

## For `GOAL_AUDIT.md`

| priority | item | what it needs |
| --- | --- | --- |
| P1 (progressed) | Source-backed CRUD adoption: price the identity fence | Change 0589 closes the DOC/PPT half. The complete-artifact fingerprint designed by 0100/0105/0119 and read-coalesced by 0143 was never priced in CPU; it is now, and its value-identical duplicate is gone. What remains open is the **write** side: the commit and save paths gain the same halving on a no-op publication and are covered only by unit tests, because no `perf-baseline` selector opens a source-backed DOC or PPT path at all. Registering such a selector is the prerequisite for any further DOC-1 work, and for any attributable measurement of the DOC read path now that hashing no longer dominates its profile. |

Supporting note for the audit body: change 0589 demonstrates the measurement
discipline GOAL step 1 asks for — the saving is proven to be *unnecessary work*
rather than relocated work, because the deterministic read counts (`read_calls`,
`read_bytes`, `len_calls`, `version_calls` per operation) are byte-for-byte
identical between the two legs on all 38 fixtures, while the hash-pass count
halves. No allocation, RSS, syscall, cold-cache, physical-device or range-source
measurement was taken, so those rows of the Phase-1 baseline are untouched. The
largest fixture either measured path admits is 1.45 MB, so DIFAT-scale behaviour
remains unmeasured, and the host has SHA-NI, so a host without it would see a
larger relative saving that was not measured.

---

## For `REPORT.md`

### 0589 — an empty overlay hashes the artifact once

Retained, partially implemented. With no physical span the CFB overlay is the
identity, so its target digest equals its source digest by construction; the
duplicate hasher is elided in `overlay::fingerprints` and `write_validated`.
DOC generic snapshot opens fall from 12 to 6 complete SHA-256 passes and PPT
text-edit opens from 4 to 2, with complete source reads unchanged. Native cycles
fall by a median 36.6% across all 38 admitted fixtures (range 25.5–48.0%, every
fixture improving); paired timing p50 improves 43.4–48.1% on the four largest,
against an A/A floor of p50 ±0.1% and p99 ≤ 2.5%. Value identity is proven by an
87-artifact differential whose two reports — every plan fingerprint, span count,
published SHA-256 and typed refusal, across empty, exact-no-op and effective span
shapes — are byte-identical. All gates pass (77 test binaries, zero failures),
including the two consumer crates. `OverlayOperationShape` now reports a no-op
plan's logical bytes hashed truthfully; its pass and chunk counts are unchanged.
The DOC open's third identity pass and the duplicate CFB index parse are
read-twice-compare defences: both are designed with a priced saving and
deliberately not implemented. No latency, allocation, RSS or cold-cache claim is
registered. OLE2/OOXML remain active; ODF is deferred until completion and iWork
excluded. [Change and limitations](0589-ole2-snapshot-fingerprint-passes.md);
[retained evidence](results/change-0589/README.md); `performance_claim: none`.

---

## For `ADR_COMPLIANCE.md`

### 0589 — the source-identity fence, priced and left intact

ADR 0006's source-identity fence and ADR 0005's `SourceChanged` contract are
unchanged. The record opens with a frozen classification of every complete-
artifact hash pass on the DOC and PPT source-backed path: identity capture
(P1, P5, P10), read-twice-compare pairs (P2, P6/P7, P8/P9), staging-window
brackets (P3), the emission proof (P4), the commit fence (P11), the composed-CFB
reopen (I3), and the ordering-critical second index parse (I2). Only P0 — the
second hasher when the span list is empty — is removed, because `apply_spans`
over an empty slice cannot write a byte and the two digests were equal by
construction. Error identity is preserved: the source comparison still precedes
the target comparison in `write_validated`, so a divergence still yields
`SourceFingerprintChanged` rather than `TargetFingerprintChanged`. ADR 0003's
rule that fingerprints are diagnostic and that exact byte equality authorizes
application is untouched; the 87-artifact differential shows every retained
digest and every published artifact is unchanged. New tests sweep **every read
ordinal** of a DOC open, a DOC readback and a PPT open with a stable-token
mutation and assert the exact number of typed refusals, so a lost fence point
fails a test rather than silently shrinking a count. The one public surface
touched is the content-free `OverlayOperationShape`, whose documented "logical
bytes hashed" now reports `source_bytes` once for a no-op plan; `is_noop()` was
already public, so no new information is exposed, and pass and chunk counts are
unchanged. No `unsafe`, no weakened limit or malformed-input defence, no new
dependency, no new public item, no ambient I/O, no Rayon pool, and no archive
type, lock or executor leaked. Change 0582's differential harness does not reach
this code (it is a `soapberry-zip` strict-layout harness); an equivalent
fingerprint differential was built and retained instead.
[Change 0589](0589-ole2-snapshot-fingerprint-passes.md);
`performance_claim: none`.
