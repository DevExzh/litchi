# ODS discrete-math performance evidence

This directory retains the reproducible process profile for the eleven ODS
discrete-math functions `COMBIN`, `COMBINA`, `FACT`, `FACTDOUBLE`, `GCD`,
`LCM`, `MULTINOMIAL`, `EVEN`, `ODD`, `DELTA`, and `GESTEP`.

The bounded matrix has 51 candidate cases (102 candidate phase groups): 11
matched controls, one scalar and one 4×4 literal-array row for each new
function, three streamed-reference rows for each of `GCD`, `LCM`, and
`MULTINOMIAL`, and three nested projected-reducer rows for each of those
reducers.  The baseline lane is commit `782339a2c` and captures only the 11
matched controls (22 baseline phase groups).  No unsupported baseline refusal
is timed as a regression comparison.

Every group has three warmups inside each child and fifteen measured fresh
child processes in both `evaluate` and `parse-evaluate` phases.  Each receipt
records elapsed time, allocator calls and bytes, live-byte balance, peak live
bytes, evaluator work, retained execution-budget memory, provider cell reads,
the result checksum, and maximum RSS from `/usr/bin/time -v`.  The runner
hashes the authored profile inputs, harness lock, source closure, toolchain,
binary, and source manifests before and after each lane, and removes its
detached baseline worktree and external Cargo target.

The scalar fixture includes a finite large `COMBIN` case (`COMBIN(1028;514)`)
whose expected binary64 bits come from the retained exact numeric golden, a
finite `COMBINA(1000;2)` lane, large integer `GCD`, and `LCM(1e308;3;0)` whose
overflowing factors are followed by an absorbing zero.  Reducer reference
fixtures are integral and sparse so all four reference sizes remain finite
while provider reads scale.  Nested rows use
`SUM(IF(range+1;REDUCER(range+0;other);0))` to expose repeated projected
branch work without changing the integer reducer fixture.

The harness is intentionally separate from production adapters.  It has no
workbook I/O, save/recalculation, native-office, networking, or cross-platform
RSS guarantee.  Final baseline and candidate captures must wait for the root
agent's recorded quiet source-freeze window; build and control preflight are
allowed before that window.
