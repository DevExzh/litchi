# 0362ute — ODG effective style dependency projection

This successor does not amend the selector-effect source line.  The
implementation is `a00a33d7c0`, whose parent is the prior receipt
`6e852ecfe7` (and whose code changes are a child of `13b4e0c2e7`).

## Effective ordered style state

Projected cross-group style checks now retain the effective final retarget for
each `(sub-edit, page)` and retire earlier retargets overwritten later in the
same ordered payload.  A final retarget back to the projected initial style is
also retired as a net no-op.  The initial and final styles are taken from the
forward projection, never inferred from untrusted inverse JSON.  Transition
events remain in the projection, so a transition that observes an intermediate
style still participates in true cross-group dependency checks.

The same effective filtering is applied to ordinary staged style changes
against joined work.  It preserves both stable sub-edit identifier orders and
ordinary-before/ordinary-after staging.  Regressions cover a page changing
`dp2 → dp1 → dp2` while another group transitions a page using `dp1`, a
same-page net-no-op style sequence against a transition, ordinary staging on
both sides of join, and a true intermediate transition that must still be
refused.

## Validation

The final source commit was validated with the memory-capped cargo shim, one
job, `CARGO_INCREMENTAL=0`, debug information disabled, offline locked
resolution, and `TMPDIR=(deleted-coord-artifact)`.
The isolated target was
`(deleted-coord-artifact)`, measured at
415 MiB after the run.  External receipts are intentionally not tracked:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg -j 1` | `(deleted-coord-artifact)/w-final-check.log` | 162 bytes | `271ef2fb0f367ffc32b5be5518d51141f815fae80f73567a8b700486620e32d5` |
| `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -j 1 -- -D warnings` | `(deleted-coord-artifact)/w-final-clippy.log` | 162 bytes | `90bb7fffb12e2f16f5785bea59c9df18779ce3372c69e7676f5ebd7cef7b5d85` |
| `cargo test --offline --locked --all-features -p litchi-odg --test package_semantics --test transition_review --test capability_transactions -j 1` | `(deleted-coord-artifact)/w-final-tests.log` | 5152 bytes | `74cd2a6333cc50cece8e69bcb60031fdd3a3d793cbe6b65aac9053bdc032e2a7` |

The final focused run passed 8 capability tests, 26 package-semantic tests,
and 22 transition-review tests (56 total).  The standalone transition rerun
also passed all 22 tests (`v-transition-tests.log`, 2090 bytes,
SHA-256 `5716ddeba23972b73c2b0be7d2c35ec949d248b66cd976e8e98d8556619e5f48`).
`cargo fmt -p litchi-odg -- --check`, `git diff --check`, the checked API
glossary policy parse, and the repository artifact guards passed.  No
benchmark, fixture, binary, or large repository evidence artifact was added.
