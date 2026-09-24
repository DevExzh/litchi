# Rejected PPTX InkAction timing attempt

This is a retained, rejected partial timing attempt from clean checkout
`260eafd391237d49373ad0ea1d73bd9348ab6226`. It is preserved for provenance
and interference review. It is not admissible timing evidence and makes no
performance claim.

The run used source checkout
`/var/tmp/pptx-ink-actions-profile-260eafd39-src.Q436KY`, external results
`/var/tmp/pptx-ink-actions-profile-260eafd39.wTYD89/results`, and external
compiled target `/var/tmp/pptx-ink-actions-profile-260eafd39.wTYD89/target`.
The target is deliberately excluded from this Git retention directory and
remains external for root audit.

The operator completed preflight, release build, host probe, and correctness
matrix. It completed 13 timed receipts through
`lane-package_read_large_shared-p1`; `lane-package_read_large_shared-p2`
started and was interrupted. A separate controlled Cargo/compiler workload was
observed during the run. Its ownership and exact phase overlap are uncertain,
so all timing output from this attempt is rejected. The observed commands and
PIDs are recorded in `operator/rejection-overlap.json`; the raw operator
command log is in `raw-results/commands.jsonl`.

The retained raw result set contains 60 files and 3,640,317 bytes. It includes
metadata, source manifest, binary hash, host/matrix receipts, commands, all
completed/interrupted lane receipts, and command stderr/stdout files. The
external target inventory at stop was 1,016 files and 357,023,074 bytes.
