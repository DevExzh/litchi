# 0699 — refusal-path attribution after rejected namespace borrowing

performance_claim: none

This is a diagnostic comparison, not a new production optimization. The main
checkout stays at the production state recorded in baseline.json. The candidate
is exactly the rejected codec witness from 0698, built only in a detached sparse
worktree. The probe adds isolated-case and profile commands; therefore these
are new binary layouts, not a replay of the original frozen 0698 executable.
The earlier regression and rejection remain valid records.

## Reproduction

Create a disposable worktree at baseline.json's baseline_head in the recorded
sibling location litchi-0699-work, using `git worktree add --detach --no-checkout`,
`git sparse-checkout set --cone crates .cargo`, and `git checkout --detach HEAD`.
Use the recorded Rust toolchain and standalone probe lock. The unchanged
root lock is retained at ../change-0698/workspace-Cargo.lock; copy it to the
disposable main checkout root before reproduction.
The packet retains source, dependency and environment hashes. Paths and CPU 12
are machine-specific and must be recorded if adapted.

Run build.py baseline followed by build.py candidate. The first copies the
probe into the isolated tree. The second applies only the saved candidate codec
there, freezes its binary, then restores that worktree's codec. Both leave the
main production tree untouched. Run cli-checks.py, measure.py, summarize.py and
profile.py sequentially, without overlapping Cargo or other performance work.
Run symbols.py for static tables, annotate.py for the retained profile
instructions and profile-summary.py for counter/loop summaries.
Then run gates.py. Generate script-hashes.json over the complete Python census,
run audit.py, cleanup.py --apply, and seal.py for the final audit/documentation checks.

The initial locked build failed before compilation because the copied lockfile
still named the previous probe package. Only that package name was corrected;
no dependency was refreshed. The first failure log is retained.

## Scope

Single-case timings use 12 process legs in ABBA, BAAB, ABBA order and three cases:
early-name-error, small-valid and late-root-error-mce. Case order reverses on
alternate legs. Each process records 300 samples after ten warmups. A separate
four-leg ABBA full-matrix control retains all ten cases in their original order.
All fixtures are authored before timing, including those not selected. Only
opened_presentation is timed; result checks, formatting and snapshot drop are
outside the timer. Exact input graph and error identities bind every run.

Profiles use a different denominator: capture plus result inspection, error
formatting, result destruction and a running assertion checksum. Setup is also
present in whole-process profiles. Three repeated (11,000 minus 1,000)/10,000
counter slopes reduce fixed startup contribution without establishing a pure
capture or cold-cache estimate. Callgraphs include an MCE-marked positive
control; absence of samples alone cannot prove a function is unreachable.
No general allocation, RSS, scaling, full-save or throughput claim follows.
