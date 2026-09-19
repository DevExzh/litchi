# ODS formula elementary-math performance evidence

This directory contains the reproducible performance profile for the bounded
ODS OpenFormula elementary-math evaluator implementation. The profile compares
the committed `8ef0057e5` baseline with one later frozen candidate using the
same process harness and immutable fixtures. The baseline does not implement
the eleven new elementary-math functions. The baseline measures matched
arithmetic/ROUND/trigonometric controls; the candidate measures all named
scalar, array, and synthetic local-reference cases. The paired controls
(`ROUND`, arithmetic, trigonometric, and the same scalar/value/reference entry
points) are the valid before/after evidence for dispatch and evaluator
overhead.

No baseline elementary-function refusal lane is captured. The baseline contains
the 15 matched controls (arithmetic, ROUND, and SIN across scalar, array, and
reference entry points); the candidate contains 42 cases: all eleven scalar
elementary functions, five representative elementary array kernels at 4x4 and
16x16, and six elementary local-reference cases alongside those controls.

The profile uses three warmup evaluations inside each child and fifteen
measured fresh child processes per case and phase.  Each child performs one
untimed direct `f64` oracle check independent of evaluator dispatch, then a
fixed repeated batch.  The oracle checks projection consistency and is not a
separate libm-accuracy implementation.  The retained receipts include elapsed samples,
p50/p95/p99 summaries, allocator calls, requested and released bytes, live
allocation balance, peak live bytes, execution work, retained execution-budget
memory, binary/source/toolchain/lockfile hashes, and maximum RSS from
`/usr/bin/time -v`.

The authored workload is intentionally bounded: scalar literals and local
references supplied by a synthetic immutable resolver, 4x4 and 16x16 literal
arrays with five representative unary elementary kernels (`ABS`, `EXP`, `LN`,
`SQRT`, and `SIGN`), representative `ABS`/`LN`/`SQRT`/`SIN` local references,
and finite execution limits. It does not exercise the production
worksheet adapter or a full workbook.  It has no filesystem, network, native-office,
recalculation, workbook-save, or package-archive path.  Results must be read
as observations for the named evaluator calls and host; they do not establish
complete OpenFormula conformance, native producer acceptance, full-workbook
throughput, or a language-level resident-memory bound.

The capture plan is in [PLAN.md](PLAN.md).  The harness and scripts are kept
under this directory so the exact inputs can be hashed before capture:

- `harness/` — standalone release binary with no production dependency;
- `run_profile.py` — paired baseline/candidate builder and fresh-process
  runner;
- `summarize.py` — deterministic paired statistics table;
- `verify.py` — fail-closed source, oracle, sample-count, allocator-balance,
  and receipt verifier;
- `results/` — retained raw captures and final report after the root agent's
  candidate freeze.

Do not run the candidate capture until the root agent records the final source
freeze.  The runner requires the freeze's selected-file hashes even though the
candidate overlay keeps the baseline Git HEAD.  A baseline-only smoke/build
may be performed earlier.  The runner removes its temporary detached worktree
and external Cargo targets after each capture and records cleanup receipts.
