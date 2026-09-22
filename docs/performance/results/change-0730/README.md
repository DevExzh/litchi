# 0730 bounded DOC validated-render handoff pilot

**Disposition: retained.** Read the [final report](../../0730-doc-bounded-render-handoff.md)
and [review](implementation-review.md) for the measured two-edit peak-memory
tradeoff and the limits of this two-fixture result. This packet preserves baseline/candidate source identity,
ordinary public lifecycle probes, capacity-based retention checks, raw process
outputs, failures, and the exact comparison schedule.

The main probe is copied from the qualified 0728 probe, with only its local
package identity renamed. It builds without performance diagnostics and is
identical for both source states. Only the two DOC public-format cases are
measured. The candidate-only retention probe separately investigates held
allocation capacity, explicit release, and successive edits; its recomputation
control is not the old baseline executable.

Commands are run from the repository root, with all Cargo/native operations
serialized. `build.py baseline` precedes production edits and `capture.py before`
records the initial baseline. After candidate qualification, `build.py candidate`,
source archiving, `retention.py build`, `qualify.py`, and `retention.py run`
qualify the source and capacity controls. `capture.py freeze`, `preflight.py`,
and `capture.py run` bind and collect the fixed comparison. All build and quality attempts remain in this packet.

Offline verification uses `source-guard.py`, `analyze.py`, `audit.py`, and
`negative-checks.py`. `artifact-seal.py --check` verifies the terminal packet
inventory. Binary identities survive removal of the two owned scratch roots in
`cleanup.json`. Reproduction requires restoring the desired archived source
state in an isolated checkout before building; the removed executable is never
silently substituted with a new one.

Allocation counts and peaks describe the measured allocator region. They do
not measure peak RSS, cold-cache behavior, remote I/O, concurrency, or general
producer compatibility. No iWork sources are part of this workstream.

Final quality is `quality-3`; baseline build is `build-1`, candidate is `build-5`,
and retention build is `retention-build-2`. Earlier failed/superseded attempts
remain archived. The retention checker label correction reuses all 24 existing
captures. The first synthetic preflight's stderr-schema mismatch and its scripts
are retained in `preflight-attempt-0`; the final preflight passes before actual
main measurement. No native measurement was selectively rerun.

The accepted result is 11.89–13.88% paired p50 improvement for NoHeadFoot,
with mixed FloatingPictures medians. Single-edit peaks are unchanged; two-edit
peaks rise 16.72% and 15.60%. All process statistics and raw reports remain in
`analysis.json` and `captures/`; capacity controls are in `retention-analysis.json`
and `retention-captures/`. End-of-region retained bytes are distinct from the
intermediate render capacity and process peak.

After cleanup, run with `PYTHONDONTWRITEBYTECODE=1` to avoid adding packet
scratch. `retention.py check` validates its raw lane. Terminal command receipts
and the artifact manifest bind the offline result and exact retained inventory.
