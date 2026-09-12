# 0541 public XLSX planning first-error guard

The [change record](../../changes/0541-xlsx-planning-error-order-guards.md)
explains six public integration tests and the 37 named fixture configurations.
[Coverage](coverage.json) records their scope and limitations.
[Test review](test-review.md) independently checks the error owners and
public selected-sheet boundary. No production optimization is admitted.

The [protocol](protocol.md), plan and quality plan are frozen before checks.
Each attempt contains the complete source manifest, exact changed test copies,
patch, raw stdout/stderr and artifact-bound receipts. `run.py` refuses occupied
outputs and checks source equality before and after each child. All Rust checks
run serially under the owned target and TMPDIR. `attempt-review.json` explains
the compile, fixture and formatting failures and the review-driven final revision; earlier results
are never relabeled as final passes.

`decision.json` designates the final attempt and counts. The verifier checks
that both changed files are tests, every other production/harness source equals
the original revision, patches reproduce retained sources, all command and
artifact hashes match, and final checks pass. It replays patches in temporary
`/dev/shm` directories that are removed on exit. The build target and test
scratch are separately removed before sealing. The verifier does not mutate
the evidence bundle.

Run `python3 -B verify.py --strict` for final source/receipt/cleanup/seal replay.
Use `--precleanup` while cleanup and the seal are pending. Reproduction must
use new paths and preserve each failed attempt; never overwrite this bundle.
No benchmark latency, allocation, hardware or profile comparison was run.
OLE2/OOXML remain the priority, ODF is deferred, and iWork is excluded.
