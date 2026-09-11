# Final review: 0514 evidence and readiness

This is the final audit of the 0514 rejection record against the raw pilot,
allocator, profile, control-repeat, restoration, cleanup, and replay artifacts.
It reviews the candidate evidence and does not treat the restored checkout as
the candidate. The authoritative decision remains [`decision.json`](decision.json):
the candidate is rejected, `candidate_kept` is false, `performance_claim` is
`none`, and `claim_authorized` is false.

## Admission result

The rejection is supported by the independent public guard and by the paired
allocator captures:

* Every cold same-value row regresses in the matched guard pilot. The six p50
  increases are `+64.43%` through `+84.64%` across tiny, medium, and dense-wide
  inputs, for both one-cell and one-percent edits. The no-op patch and exact
  source-byte oracles pass, so this is an admission failure caused by work and
  memory cost rather than a semantic failure.
* The dense cold one-cell allocator row grows from 265,295 to 597,095
  allocation calls and from 47,016,362 to 73,571,934 allocated bytes. Its
  incremental operation-region peak grows from 27,278,293 to 34,655,117 bytes
  (`+27.04%`). The dense one-percent row grows from 533,203 to 1,196,803
  calls, from 94,143,070 to 147,254,214 bytes, and from 39,575,534 to
  46,952,358 incremental peak bytes (`+18.64%`). The smaller shapes show the
  same direction. These values are operation-scoped allocator observations;
  they are not document peak or RSS claims.
* Each role has two allocator captures with 10 measured samples and one
  warmup. The per-scenario vectors repeat exactly within the before role and
  within the after role. The raw reports therefore support the stated
  candidate/control comparison; they do not turn an after-only capture into a
  before baseline. Normal uninstrumented allocation totals remain unavailable,
  as recorded in [`summary.json`](summary.json).

The main changed-edit pilot remains mixed: eleven rows improve by about
`0.42%` to `6.21%`, while dense-wide `xlsx_one_percent_commit` regresses
`+11.09%`. The selected three-call Callgrind profile falls `6.30%` in
instruction references, but it is changed-output profile context and cannot
offset the cold no-op gate failure. No full candidate native lane was required
after this predeclared rejection, and no alternative design is ranked.

## Control and measurement audit

The completed control repeats are consistent with the report and remain
properly separated from candidate deltas:

* Main native controls use 500 samples and five warmups per row. Same-build p50
  drift ranges from `−1.67%` to `+2.29%`, with no main drift flag. Main
  whole-child RSS is `140,232 -> 137,348 KiB` (`−2.06%`).
* Guard controls use 100 samples and three warmups per scenario. Same-build p50
  drift ranges from `−4.63%` to `+3.88%`. The tiny cold same one-cell p95 has
  one absolute drift flag at `−11.07%` against its 10% threshold. Guard
  whole-child RSS is `169,312 -> 180,284 KiB` (`+6.48%`); this is a before-role
  same-build repeat and has no candidate RSS comparison.

The guard tail and RSS observations are residual control-context caveats, not
candidate memory deltas. They do not erase the much larger paired cold
operation allocator increases that reject admission. The change report and
replay summary disclose both flags, so they must remain visible in the sealed
record rather than being averaged away.

## Custody, correctness, and ADR consistency

The candidate patch is bound to base revision `33a21e0f0`; its exact replay is
recorded as passed. [`restoration.json`](restoration.json) records exact
restoration of all 14 candidate paths and the control source manifest. The
final candidate validation receipt reports 966/966 owner tests with zero
ignored, including 17 fusion tests; the associated Clippy, formatting,
boundary, and strict-claim receipts pass. This evidence validates the tested prototype behavior and custody. It does not authorize retaining the candidate after
the performance gate failure.

The dedicated [`adr-review.md`](adr-review.md) explicitly covers the relevant
matrix: ADR 0001 priorities and API layers, 0003 snapshots/edits/patches and
concurrency, 0005 I/O/memory/performance scope, 0006 validation/security/
compatibility, 0008 migration and verification, 0011 OOXML physical package
ownership, and 0018 XLSX calculation-chain ownership. The global 0514 sections
in [`ADR_COMPLIANCE.md`](../../ADR_COMPLIANCE.md),
[`BASELINE.md`](../../BASELINE.md), [`GOAL_AUDIT.md`](../../GOAL_AUDIT.md),
[`HOTSPOTS.md`](../../HOTSPOTS.md), and [`REPORT.md`](../../REPORT.md) agree on
the rejected status, no authorized speed claim, exact restoration, and the
remaining OLE2/OOXML priority. Their detailed numbers point back to the change
report and raw bundle rather than inventing a full after-native comparison.

The post-cleanup replay in [`replay-after-cleanup.json`](replay-after-cleanup.json)
exits zero, is equivalent to the retained summary, reports no stderr, and
confirms the owned scratch is absent. [`cleanup.json`](cleanup.json) records
removal of the owned temporary worktree and scratch files while retaining the
source hashes, binaries' identities, raw reports, profiles, and logs required
for review. The verifier receipt and its summary hash agree.

## Readiness and checksum sealing

The candidate is **not production-ready and must remain rejected**. The
evidence bundle is technically ready to seal: the decision, exact restoration,
control caveats, paired no-op/allocator evidence, cleanup, and post-cleanup
replay are all present and mutually consistent. No rerun is justified by the
control drift flags, and this review found no evidence that would authorize a
performance claim.

The only packaging follow-up found by the independent audit was the absent
`SHA256SUMS` inventory referenced by README.md. Root resolves it in the final
sealing step: the inventory includes this review and the retained evidence,
and its entries are checked before commit. No experiment or source change is
needed for that packaging step.

The next optimization investigation remains OLE2/OOXML, including the open
XLSX changed-output and DOCX publication work. ODF work stays deferred until
the OLE2/OOXML goal is complete, and iWork remains excluded.

No source, build, test, benchmark, or capture command was run for this final
review.
