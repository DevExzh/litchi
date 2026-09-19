# Final capture verification notes

The corrected custody capture was run after the isolated candidate received the
missing authored `numeric_oracle.py`. No compiled Rust, frozen contract, native
receipt, golden, or profile harness input changed for this rerun.

The local fail-closed verifier passed with 510 baseline records and 3,090
candidate records across 103 candidate cases, two phases, three warmups, and
fifteen fresh measured children per case and phase. Root's independent verifier
also passed all 206 groups and 3,600 samples; it confirmed source, generator,
native, golden, and lock custody. The corrected candidate manifest records the
`numeric_oracle.py` SHA-256 and has stable selected-source, workspace-source,
profile-input, and git-head fences.

Across matched controls, root's independent p50 review found time deltas from
-3.238866% to +2.347249%, unchanged allocation metrics, and RSS deltas from
-2.516556% to +4.216074%. No matched metric crossed the +/-5% review threshold.
Candidate-only dispersion rows are absolute measurements and have no baseline
comparison. The raw child stdout, `/usr/bin/time -v` receipts, measurements,
source manifests, cleanup receipts, and summaries remain under this `results/`
directory.

Exact scratch removal paths are recorded in `scratch-cleanup.json`.
