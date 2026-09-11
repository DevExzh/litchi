# 0503: reconcile the current non-iWork performance evidence

This batch updates the goal audit, hotspot queue, and report against retained
changes through 0502. It changes no production code and makes no new speedup
claim. The previous batch made concrete capability and measurement progress;
this continuation verifies its evidence and changes the next investigation
priority to the measured ODG open regressions while preserving the richer
metadata semantics.

The [machine-readable receipt](../results/change-0503/audit.json) binds the
checked raw reports, catalogs, ODG summary, and claim registry by SHA-256.
Both 0501 full-run report/catalog pairs pass the coverage validator, the 38
coverage tests pass, and strict registry verification validates 10 claims.
The 0502 summary reproduces exactly from 16 raw reports (eight matched pairs,
3,200 measured samples), including the deterministic bootstrap intervals.
These are historical-data checks, not a new current-worktree benchmark or a
workspace correctness run. No Rust tests or builds were needed for this
documentation-only batch.

The complete 0501 final verifier cannot replay in the current environment:
it exits with `before: frozen binary is missing`. Its historical successful
receipt is preserved. A historical cleanup statement that binaries were retained
in local tmpfs does not establish their present availability. Rebuilding and
capturing a fresh matched pair remains possible; neither a missing executable
nor a passing report validator erases the retained measurements or proves
current runtime behavior.

The actionable queue retains all important adverse results: ODG plain and
metadata open regressions in 0502, small local Part-batch overhead and delayed
provider tails in 0499, and the small-selection latency and RSS flags in 0500.
The 0501 timing-report gate is satisfied by retained reports; broad CRUD,
provider, native-producer, memory, and scaling requirements remain incomplete.

No temporary build directories or binaries were created. Python ran with `-B`;
the verifier-generated replacement of its historical receipt was restored.
Only the intended documentation and compact audit receipt are committed.
