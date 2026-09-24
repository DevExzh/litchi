# Untimed preflight receipt

Before the final profile capture, build the frozen candidate harness in release
mode and run every named case with `--preflight-only` in both evaluator phases.
This executes the same parse, correctness, resolver, budget, and formula-error
checks used by a timed child while collecting no timing sample. Run the detached
d6 baseline against its 21 matched controls in both phases as well.

Each preflight line includes the case name and resolver-read count. The
projected LARGE criterion should report 26 reads: two position-sensitive
G-parameter reads plus twelve reads for each of its LARGE and COUNTIF ranges.

Record the commands, exit codes, case counts, candidate and harness hashes, and
the exact resolver-read observations for the projected MUNIT criterion in this
directory. The final receipt must be written before the capture starts, because
the profile input hash covers all files in the performance directory. Release
target directories and temporary baseline worktrees must be removed; the
capture script records their exact paths in `results/*/target-cleanup.json` and
`results/capture-summary.json`.
