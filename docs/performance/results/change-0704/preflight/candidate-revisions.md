# Candidate preparation history

The isolated coder handoff was reviewed before measurement. Root integrated
it, corrected final-table accounting and test plumbing, and replaced repeated
capture-local scans with append, checked running charge, and final sort/dedup.
The first built candidate was not timed. Its source diff and build receipts
are retained in `candidate-01/`; those executables were superseded.

The first focused run compiled and passed 22 of 24 tests. Two test mistakes
were corrected: the observation counter included two test-hook source reads,
and an invalid-root fixture changed only its opening tag. The original log,
source census and test source remain in `functional-01/`.

The first Clippy run found two expression-style violations. The log and source
are in `clippy-01/`. Both were fixed. Later tests added owner-only rebind
projection, equal-content allocation misses and deterministic final-table
reservation refusal. The final focused inventory contains 26 passing tests.

The first retention observer stopped at its 64-shape per-slide limit on a
108-shape slide. Its source and stderr are preserved in
`candidate-01/retention-probe/`. The diagnostic ceiling was raised to 1,024;
the final observer was rebuilt against the final candidate and both sources
passed. No native timing sample came from the failed observer or earlier
candidate source. Source-bound final build receipts are authoritative.
