# Verification disposition

The frame-pointer capture, symbol identity check, stack analysis, and portable
stack-analysis mutation checks pass. Both retained reports pass the bound
report verifier: each has 100 retained samples, 500 checked cache deltas and
1,200 checked read deltas. Their source manifest covers the same 6,634
Rust/TOML/lock files and their executable hash is identical.

Full portable bundle replay passes before and after removal of the temporary
executable. It invokes both report validators and reruns stack derivation,
checks cross-provider output identity, and rejects five mutation classes:
numeric API sums, output size, the bound report verifier, decoded stack bytes,
and summary counts. Numeric/output mutations also exercise the report
validator directly, independently of their stale artifact hashes. The
separate stack-analysis portable check rejects protocol, index, receipt,
source-audit receipt and retained-stack custody mutations.

The first two full replay attempts failed because of checker assumptions;
their exact scripts, receipts and logs remain classified as failed evidence.
The [correction record](validation-corrections/README.md) explains both fixes.
The successful pre-cleanup attempt is `portable-before-cleanup-v3`; the
successful post-cleanup attempt is `portable-after-cleanup`.

`replay-copy.py` exports a sealed bundle before creating the original bundle's
running receipt. Its export records retain the input inventory and verifier
hashes and confirm temporary-directory removal. Thus each replay validates
one complete prior seal; the final inventory additionally includes all
terminal replay receipts and logs.

The [source review](checks/source-review.md) accepts the attribution scope
subject to the final seal. No Rust changes or new Rust test results are
claimed. Historical strict lint debts remain open, and the
[README](README.md) states the non-hermetic rebuild, machine provenance,
symbolization and sampling limits.

The task's copied executable directory is absent. `cleanup.json` confirms the
original executable and both target directories remain and pins the preserved
`docs/GOAL.md` hash. All copied replay directories are removed. No global goal
completion or measured optimization is claimed.
