# ODS discrete-math evaluator performance plan

The profile compares candidate implementation work against matched controls
from committed baseline `782339a2c`.  The baseline predates the eleven named
functions, so its lane contains arithmetic, `SIN`, `IMSUM`, `DSUM`, and `SUM`
controls at the same scalar, literal-array, and streamed-reference entry
points.  Unsupported new-function results are capability evidence only and
are excluded from timing comparisons.

The initial workload is deliberately bounded at 51 candidate cases (102 phase
groups, plus 22 baseline control phase groups):

* 11 matched controls cover scalar arithmetic, `SIN`, `IMSUM`, `DSUM`, and
  `SUM`, plus 4×4 literal and 16×4 reference arithmetic/`SIN`/`SUM` paths;
* each new function has one scalar row and one 4×4 literal-array row;
* `GCD`, `LCM`, and `MULTINOMIAL` each have 64×4, 256×4, and 1024×4
  streamed local-reference rows;
* each reducer has 64×4, 256×4, and 1024×4 nested projection rows using a
  literal arithmetic expression over one reference and a second reference;
* the scalar cases exercise finite large `COMBIN`, large `GCD`, and
  overflow-then-zero `LCM` vectors.

Each case runs in a fresh release child with three untimed warmups and fifteen
measured samples.  The `evaluate` phase reuses a parsed expression; the
`parse-evaluate` phase parses inside the timed batch.  Fixed repeat counts keep
short scalar calls measurable while large reference matrices are evaluated
once per sample.  A borrowing in-memory resolver counts every `read_cell`
call.  Allocator counters are reset at the start of each measured child and
the verifier requires requested/released bytes and live bytes to balance.

Before either final lane, freeze the selected candidate files and authored
profile inputs, then record the full workspace source closure and lock hashes.
Run baseline from a temporary detached worktree at `782339a2c`, followed by
the frozen candidate with identical harness, lockfile, toolchain, phases,
warmups, samples, and input hashes.  Retain raw JSON, `/usr/bin/time -v`
receipts, build logs, source manifests, cleanup receipts, and verifier output.

The root agent should expand the matrix only when a contract or reviewer
finding requires a named additional row.  Any matched-control time, allocation,
RSS, budget, or provider-read change above five percent receives scoped review
before it is described as a regression.
