# ODS formula trigonometry performance evidence

This directory contains the reproducible performance profile for the bounded
ODS OpenFormula trigonometric evaluator implementation.  The profile compares
the committed `069c65870` baseline with one later frozen candidate using the
same process harness and immutable fixtures.  The baseline does not implement
the new trigonometric functions.  The baseline measures the ten matched
arithmetic/ROUND controls; the candidate measures all fifty named scalar,
array, and synthetic local-reference cases.  The paired controls (`ROUND`,
arithmetic, and the same scalar/value/reference entry points) are the valid
before/after evidence for dispatch and evaluator overhead.

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
arrays with five representative unary trigonometric kernels (`SIN`, `COS`,
`ASIN`, `ACOSH`, and `TANH`), representative `SIN`/`ACOSH`/`ATANH` local
references, and finite execution limits.  It does not exercise the production
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
