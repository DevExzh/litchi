# ODS financial performance plan

This is a pre-capture plan for the twenty-function financial slice. It defines
the workload, comparison boundary, and evidence custody; it contains no timing
result and records no build or performance execution.

The normative input is [`../contract.md`](../contract.md), whose current
contract snapshot is SHA-256
`fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b`. The
accepted numerical profile is recorded in
[`../numerical-design.md`](../numerical-design.md). The ODF archive and Part 4
HTML hashes remain the values recorded by that contract. The comparison
baseline is the prior date/time commit
`f945a5129ad8967a863cf4937de01ee2152b9fbf`; the financial candidate identity
must come from a later source freeze and a checked `candidate-freeze.json`.
The ambient workspace lock is never a substitute for the isolated lock used
by a capture.

## Comparison boundary

The baseline and candidate are separate isolated checkouts. The baseline is
checked out at `f945a5129a` and the candidate is checked out at the final
financial source freeze. The runner records the complete commit IDs rather
than relying on the short labels above.

The earlier commit contains financial names in some registries and in the
legacy evaluator, but a registered name is not evidence that the target ODS
evaluator supports the function. A capability preflight must classify every
financial row on the baseline as `supported`, `unsupported`, formula error,
typed failure, or harness failure. An `Unsupported` result for a financial
row is a capability observation. It is not a zero-cost measurement and it is
never used to claim a candidate speedup or a before/after regression. The
candidate must pass the same row's independent expected-value, shape, and
resource checks before that row is timed.

Only controls that produce the same exact semantic result and declared
accounting on both revisions are valid matched comparisons. The financial
rows are candidate-only whenever the baseline lacks their ODS implementation;
their report gives absolute latency, work, reads, storage, allocation, and
scaling by input size. It does not attach a baseline delta or a speedup to
those rows. If a financial row happens to pass on both revisions, it becomes a
matched row only after both capability receipts and the exact contract checks
agree.

The matched control set is copied from the reviewed predecessor performance
workflow and frozen by formula, evaluation path, shape, and fixture. It must
include small scalar arithmetic and trigonometric controls, rectangular array
arithmetic and trigonometric controls, a scalar and a reference aggregate,
one conditional/reference aggregate, and lazy `IF`/`IFERROR`/`IFNA` controls.
At least one control uses a scalar result and at least one uses a matrix
result. These controls measure regressions in the evaluator bridge, matrix
scheduler, resolver, allocator, and error path that financial dispatch could
otherwise hide. A financial function is never used as an “unchanged” control.

## Function scope and workload matrix

The authoritative `case-matrix.json` will contain every row name, exact
formula, evaluation path, execution geometry, expected outcome, reference-read
bound, and fixed repeat count. Its declarations are the input to the runner
and analyzer; hand-maintained case lists in a report are not sufficient.

| group | functions | workload purpose |
| --- | --- | --- |
| scalar-kernel | `EFFECT`, `FV`, `IPMT`, `ISPMT`, `NOMINAL`, `NPER`, `PDURATION`, `PMT`, `PPMT`, `PV`, `RRI` | zero-rate and nonzero-rate arithmetic, optional controls, stable `ln1p`/`expm1` branches, checked finite results, and domain errors |
| sequence-reducer | `CUMIPMT`, `CUMPRINC`, `FVSCHEDULE`, `MIRR`, `NPV`, `XNPV` | ordered sequence scans, period loops, product/sum state, paired value/date traversal, array and reference storage |
| root-solver | `IRR`, `RATE`, `XIRR` | bracket expansion, transformed positive and signed negative branches, repeated residual work, cancellation, convergence, and no-root behavior |
| matched-controls | unchanged scalar, array, aggregate, conditional, and lazy evaluator rows | regression detection in shared parsing, scheduling, resolver, cache, allocation, and publication paths |

Every function has a small literal core row and at least one constraint or
failure row. The scalar-kernel rows additionally cover zero and nonzero rates,
ordinary and beginning-payment controls where applicable, small and large
finite magnitudes, near-zero stable branches, overflow/underflow refusal, and
the exact integer-versus-number conversion profile. `RRI` includes positive
ratios, two negative inputs, the exact binary64 reciprocal-integer negative
ratio, and the required non-integral reciprocal refusal. `EFFECT` and
`NOMINAL` include reciprocal pairs so a faster result cannot be accepted when
the compounding relation changed.

Sequence reducers use source sizes `N = 4, 64, 1024, 4096` where the
function admits a sequence. The largest row uses a lower fixed repeat count,
recorded in the matrix, instead of changing the workload after observing its
latency.

* `FVSCHEDULE` uses ordered schedule references and arrays, including a zero
  factor, mixed signs, and a late formula Error. It records one physical read
  and one sequence work charge per inspected cell, including a skipped cell.
* `NPV` uses one reference, multiple ordered `ReferenceList` areas, and split
  arguments. Area order, sheet order, row-major order, one-based discount
  indices, and a late typed failure are observable. Its expected reads are the
  sum of all admitted areas, not the number of output coordinates.
* `CUMIPMT` and `CUMPRINC` use period spans `4, 64, 512, 2048`. Their matrix
  records period work and checked integer range handling; they have no hidden
  resolver reads. Both ordinary and beginning-payment types are represented,
  including the deliberate `CUMPRINC`/`PPMT` boundary profile.
* `MIRR` uses rectangular arrays of `4, 64, 1024` elements with Text, Empty,
  Logical, positive, and negative positions. Positive and negative masks share
  the original positions; the two signs are not compacted into separate
  period sequences. Direct scalar Reference and known ReferenceList refusal
  rows assert the selected zero-read shape behavior.
* `XNPV` uses values and dates with `4, 64, 1024` admitted elements, including
  different rectangular geometries with equal flattened counts, a values-first
  scan followed by a dates scan, a nonannual interval, and a late date-side
  failure. It also has the `rate <= -1`, unequal-count, non-number, and date
  ordering rows.

The root-solver matrix keeps the sequence length separate from iteration cost.
Each solver has a short vector and a longer vector, a canonical convergent
root, a near-zero root, a supplied-guess variant, a no-root or nonconvergent
case, and a multiple-root or branch-selection case where the contract defines
one. `IRR` covers a root below `-1` selected by a below-`-1` guess as well as
the ordinary positive branch. `RATE` covers positive and negative rates,
integral `Nper` on the signed branch, fractional `Nper` refusal, and the
explicit `rate == -1` boundary. `XIRR` covers nonuniform dates, a negative
rate above `-1`, unequal/masked pairs, and a date-weighted no-root case; it
does not use the negative-base branch because its exponents can be fractional.

Projected rows wrap scalar controls and complete sequence descriptors in a
small matrix `IF`, such as a two-coordinate `NPV`, `FVSCHEDULE`, or `XNPV`.
They verify that complete sequence inputs remain complete while rate, payment,
guess, type, and pay-type controls remain position-sensitive. A nested scalar
criterion such as `SUM(MUNIT(...))` is retained as a separate position-
sensitive row and is not used to justify invariant cache reuse. The same
formula is run through both the pre-parsed evaluation and parse-evaluate
paths.

## Resource and failure lanes

Resource rows are measured as their own case class and are not silently
converted into formula errors. For each streaming reference size, the matrix
contains an exact-limit and one-under-limit `max_reference_cells` row. The
reference limit covers admitted geometry and cumulative successful reads. The
work limit is tested at the exact required charge and one below it; root rows
also test the term, derivative, iteration, and total evaluation caps. Numeric
scratch and result arrays are tested at exact storage capacity and one over
capacity with fallible reservation and drop receipts.

Cancellation rows cover cancellation before the first read, after a physical
read, during a reducer scan, and during a root iteration. Sticky cancellation
is repeated four times with the same child contract; a child that was cancelled
after one successful read must report one total successful read across those
four repeats, not four reads. Source-version changes are tested both after a
complete sequence and immediately before publication. A late provider,
allocation, resource, cancellation, or source failure must supersede a
retained formula Error. `IFERROR` and `IFNA` may handle formula-error values,
but never typed evaluator failures.

Known wrong sequence shapes are preflighted before the resolver: a statically
known `ReferenceList` refusal for `IRR`, `FVSCHEDULE`, `XIRR`, or `XNPV` has
zero cell reads. Computed children may run far enough to establish their
runtime type, so their read bound is case-specific. The matrix distinguishes
zero-read shape refusal from a provider failure after partial reads. Every
inspected physical cell charges work and checks the reference limit and
cancellation before and after the resolver, even when conversion later skips
Empty, Text, or a formula Error. Borrowed resolver text remains borrowed while
conversion is charged.

Root rows record the admitted numeric scratch capacity and release order. A
root scan reads the source once and repeated residual evaluations use only the
bounded converted buffer; they must not reread a provider. The matrix records
term, derivative, iteration, and total-evaluation work separately from
resolver reads. The accepted numerical profile is 64 bracket expansions, 128
solver iterations, and 256 total residual/derivative evaluations, with the
caller work budget still able to fail earlier. No hidden clock, random source,
filesystem, network, or external rate provider is permitted.

### Coverage reconciliation

The matrix has eleven matched controls: scalar arithmetic, scalar
trigonometry, scalar `SUM`, rectangular arithmetic and trigonometry, reference
`SUM`, reference `AVERAGE`, conditional reference `SUMIF`, lazy `IF`,
`IFERROR`, and `IFNA`. The conditional row uses a 64-cell criterion and
selected range and records the selected-cell read count separately from the
criterion scan. This keeps the aggregate, conditional, and error-path
controls promised above in the same fixture and comparison boundary.

The candidate side has a literal core row for each of the twenty contracted
functions, period-size rows for `CUMIPMT`/`CUMPRINC` at 4, 64, 512, and 2048,
zero-rate scalar branches, negative-input `RRI`, zero and formula-error
`FVSCHEDULE`, split and ordered-list `NPV`, a projected `NPV` branch, the
negative-base `IRR` branch, a non-annual `XIRR`/`XNPV` pair, three
sequence-list refusals, domain-error rows, and a late date-side provider
failure. The existing 4/64/1024/4096 streaming rows and 4/64/1024 root/array
rows remain unchanged. Matrix and harness case names, formulas, outcomes,
shapes, and read bounds are kept in lockstep.

The 1024/4096 `NPV`, 1024 `XNPV`, and 1024 `MIRR` rows carry an explicit
`max_steps = 8,000,000` evaluator budget in both the harness and matrix. Their
larger admitted sequences legitimately exceed the default one-million-step
profile; the per-row budget makes that workload choice visible and keeps it
separate from the intentionally small work-limit refusal row.

Some semantic requirements are intentionally outside this process-level
profile. Source-version mutation, allocation-failure injection,
fallible-scratch reservation/drop-order receipts, and native recalculation
provenance need the focused evaluator/resource/native gates because the public
profile resolver has no source mutator or allocator-fault hook. The capture
still retains cancellation, reference/work limits, formula-error retention,
typed provider failure, zero-read shape refusals, and exact read accounting;
the focused gates own the remaining failure injections. These exclusions are
documented here so the timing matrix does not imply that a process fixture
proved a semantic gate it cannot exercise.

## Independent preflight

The independent financial oracle is prepared outside the timer. It derives
finite scalar and matrix outcomes from the contract equations and the accepted
numerical profile; it never imports the evaluator or calls a spreadsheet host.
Expected values, formula-error identity, typed failure kind, output shape,
execution geometry, exact formula string, evaluation path, and read/work
contract are emitted in the preflight receipt. Financial roots additionally
record the selected branch, residual tolerance, root-selection tie rule, and
convergence classification. An `AnyNumber` expectation is not sufficient for
a financial case.

The runner compares the receipt with `case-matrix.json` before timing. It
checks the full scalar or numeric matrix payload recursively, distinguishes
booleans from numbers, checks exact error/failure strings or typed kinds, and
checks the declared execution geometry independently of the result payload.
All heavy oracle work, date conversion, expected arrays, and large reference
fixtures are prepared before the timed child. The timed path performs only the
evaluation, prepared-result validation, complete-result checksum, and drop.
This catches a stale expected value or a missing array checksum without adding
unbounded oracle work to the measurement.

The baseline preflight runs all matched controls and all twenty financial rows.
Rows that are unsupported or otherwise unavailable on the baseline are retained
as capability receipts and excluded from deltas. A baseline formula-error row
is comparable only when the candidate produces the same contract error with
the same shape and accounting; an accidental unsupported/error path is not a
control.

## Measurements and analysis

The capture uses the established two-phase protocol: `evaluate` on a
pre-parsed expression and `parse-evaluate` from formula text. The default
sample geometry is three warmups and fifteen fresh child samples, with the
fixed per-case repeat count declared in the matrix. Scalar rows may use a
larger fixed repeat count, while large references and failure rows use a
smaller one; both revisions use exactly the same count. No sample or case is
dropped after seeing its latency.

Each raw sample retains monotonic `elapsed_ns`, exact floating-point elapsed
time normalized by repeat, allocator calls and requested/released/live/peak
bytes, retained evaluator storage, work, successful resolver reads, input and
output bytes, checksum, result shape, and external RSS/HWM when available.
Root rows also retain iteration and residual counters. Report p50, p95, p99,
mean, minimum, maximum, and sample count for every phase and every input size.
Read/work/storage/allocation changes are reported independently; equal reads or
equal allocator counters do not establish a causal optimization.

For valid matched controls, pair samples only within the same case, phase,
repeat count, and input geometry. Report the signed median delta and the
individual rows rather than only a geometric mean. Use a fixed seeded bootstrap
(seed `20260920`, 100,000 paired resamples, 95% interval) over the exact
normalized elapsed values for descriptive uncertainty. A row at least 5% slower
or faster in latency, or at least 5% different in peak RSS, is a review trigger
and reports absolute values, signed delta, interval, sample count, and host
load. It is not dismissed as noise and it does not establish a cause or justify
a favorable rerun. Material scaling loss is also a trigger even below 5% at a
single size.

Candidate-only financial rows report absolute cost and scaling curves. For
streaming functions, check that reads and work grow with admitted input size
while state stays within the declared bound. For root functions, report
iteration/evaluation distributions and cancellation behavior separately from
sequence-read cost. The report must not manufacture a baseline delta for an
unsupported function.

## Source, build, and capture custody

Before capture, the source closure must include the evaluator dispatch and
financial scalar/value modules, shared `numerics.rs`/`dyadic.rs` helpers, date
conversion and paired/reference scheduling used by XIRR/XNPV, matrix/cache
code used by projected financial calls, relevant core budget/execution code,
all financial semantic/limits/oracle tests, the feature matrix, contract,
numerical-design review, independent oracle and vectors, native provenance,
the harness, and the isolated `Cargo.lock`. Recursive globs must include any
future `financial`, `cashflow`, `solver`, `value`, or numerical child modules;
the ambient root lock and unrelated untracked files are excluded.

The freeze manifest hashes every selected source and documentation input and
records the contract, oracle, matrix, harness, lock, baseline commit, candidate
commit, Rust toolchain, target, allocator, compiler flags, CPU count, and
allowlisted process-load snapshot. It records binary identities before and
after measurement, but it need not retain build targets or claim that binaries
are archived. Baseline samples are captured before candidate samples, and the
chronology is retained with the host caveat that CPU affinity and load are
observations rather than an isolation guarantee.

The result directory retains the frozen matrix, baseline capability receipt,
candidate and baseline preflight receipts, raw measurement JSONL, external
`/usr/bin/time` receipts, environment/context snapshots, build logs, summaries,
analysis, and a flat `retained-files.json` mapping every retained relative
path to its SHA-256. A failed setup, compile, preflight, or verifier attempt is
moved to a diagnostic directory and described; it is never overwritten by a
metric-selection rerun. Temporary targets, staged checkouts, and registries
are removed only after their hashes and logs are retained.

This plan establishes the financial workload and evidence boundary only. It
makes no implementation, correctness, speedup, regression, allocation, RSS,
throughput, or scaling claim, and it does not authorize Cargo or performance
execution before an explicit source-freeze handoff.
