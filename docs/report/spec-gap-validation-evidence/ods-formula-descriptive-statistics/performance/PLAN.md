# Descriptive-statistics performance profile plan

**Protocol:** run only with root's frozen candidate manifest and quiet-window
authorization. Capture outputs are written under `results/`; once invocation
starts, the source closure, contract, harness, and profile inputs are immutable.

This profile covers `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`,
and `SKEWP` against baseline commit
`b8e5d5fe257fd95747c69a3c44a53cedd96f77ed`. The baseline is checked out in a
detached worktree and uses the retained lock at
`docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/gates/Cargo.lock`.
The descriptive harness carries a copy of that lock, so the two lanes use the
same resolved dependency graph.

The eventual capture will use three untimed warmups and fifteen fresh child
processes for every case in both `evaluate` and `parse-evaluate`. The child
records elapsed time, allocator calls and requested/released bytes, live-byte
balance and peak live bytes, evaluator work, retained budget memory, resolver
reads, checksum, and external `/usr/bin/time -v` peak RSS. A source/profile
fence is checked before and after each lane. No result will be retained as a
performance claim unless its correctness preflight, source closure, lock,
toolchain, and cleanup receipts pass.

## Case matrix

The matched baseline controls are the 21 existing rows below plus the three
representative scalar rows listed after them. They provide the
same arithmetic, scalar, database, aggregate, conditional, array, and
reference controls used by the earlier reducer profiles:

```text
scalar-control-arithmetic
scalar-control-sin
scalar-control-imsum
database-control-dsum
scalar-control-average
scalar-control-counta
scalar-control-var
scalar-control-stdev
database-control-dvar
database-control-dstdev
array-control-4x4-arithmetic
array-control-4x4-sin
array-control-16x16-arithmetic
array-control-16x16-sin
reference-array-16x4-arithmetic
scalar-aggregate-sum
literal-aggregate-4x1-sum
reference-aggregate-64x4-sum
reference-conditional-256x4-sumifs
reference-control-average
reference-control-counta
```

The representative controls keep the existing ordering implementation visible
as a nearby scalar-dispatch reference and are captured on both sides:
`representative-median`, `representative-rank`, and
`representative-percentrank`.

Each descriptive function has the following bounded lane slots, for 13 slots
per function and 91 candidate slots in total:

| lane | input and purpose |
| --- | --- |
| `scalar` | One scalar call with a finite fixture. |
| `inline` | One inline array call; the fixture is small enough to expose literal-array handling without making parsing dominate every row. |
| `reference-64` | A 64x4 borrowed reference for the small scan. |
| `reference-256` | A 256x4 borrowed reference for the middle scan. |
| `reference-1024` | A 1024x4 borrowed reference for meaningful linear scaling. |
| `projected-64` | A projected 64x4 reducer whose complete reference must be preserved through the projection. |
| `projected-256` | The matching projected middle-size lane. |
| `projected-1024` | The matching projected large-size lane for cache/work scaling. |
| `domain` | The function-specific signed, non-positive, zero, or insufficient-pass-count input required by the frozen numerical contract. |
| `list-admit` / `list-refusal` | One reference-list shape row per function. `AVEDEV`, `GEOMEAN`, `HARMEAN`, `KURT`, and `SKEW` use `list-admit` with 16 cells per pass; `DEVSQ` and `SKEWP` use `list-refusal` and reject before any resolver read. |
| `error` | A reference containing a formula error after valid cells, checking retained-error precedence and continued scanning. |
| `cancellation` | A deterministic cancellation after a bounded prefix, checking the typed failure and exact partial read/work receipt. |
| `resource` | A deterministic reference-cell/work limit refusal, checking the typed failure and read-before-limit ordering. |

The harness names are formed as `{lane}-descriptive-{function}` (for the size
lanes, for example, `reference-descriptive-256-avedev`). The scalar, inline,
and reference lanes use a positive fixture shared by all seven functions so
the reference size lanes remain directly comparable; the domain rows exercise
signed and non-positive behavior where it is function-specific. The reference
fixtures contain stable duplicates and a deterministic tail so linear
input-size behavior is visible without multiplying the matrix by every
permutation. The pass-count and list-shape expectations follow the
reviewed contract and are recorded in the harness case names and read-bound
verifier. The untimed preflight records each edge, typed failure, and exact
read receipt before any sample is accepted.

The matrix deliberately keeps the projected size lane separate from direct
reference scaling. This makes a complete-reference demand-cache result
observable while keeping nested scalar parameters position-sensitive. Any
future parameter expression must be scalar at the reducer boundary; a matrix
parameter is a typed refusal case with zero resolver reads where the contract
requires preflight rejection.

## Read and resource assertions

The final verifier will require exact observed reads rather than broad timing
windows. Scalar and inline lanes must perform zero resolver reads. A complete
64x4, 256x4, and 1024x4 reference scan exposes 256, 1,024, and 4,096 cell
reads for one-pass `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, and `SKEWP`.
`AVEDEV` performs an exact-mean scan plus one charged replay and therefore
exposes 512, 2,048, and 8,192 reads. The projected invariant rows use the
complete sequence descriptor once; their reducer reads stay at those same
one-pass or two-pass counts and are never multiplied by the two output cells.
The admitted 16-cell list row follows the same rule (`32` reads for `AVEDEV`,
`16` for the other admitting functions), while each refused list row is a
zero-read shape preflight.
An exact harmonic replay is a separate adaptive/error case when the owner
adds one; it is not conflated with the ordinary one-pass lane. Formula-error
rows finish their required primary scan and skip a centered replay once the
retained error is final.

The list/domain/type refusal rows must be checked before resolver access when
the contract says the shape or type is rejected. Formula-error rows must keep
scanning admitted cells so a typed resolver/resource/cancellation/source
failure can supersede a retained formula error. Cancellation and resource
rows will record the configured threshold, successful reads, work, and failure
kind in the case manifest; a range such as “0..N” is not sufficient evidence.

All rows retain the allocator and budget evidence needed to check bounded
state. A reducer may use fixed-size accumulators, but the profile will report
peak live bytes and retained budget memory so an implementation that grows a
cell vector is visible. Borrowed resolver text and source-version/cancellation
fences remain part of the correctness gate.

## Source closure and readiness

The runner records exact SHA-256 values for the root manifests, the selected
evaluator/value/statistical/descriptive Rust files, including recursive
`descriptive/**/*.rs` and `value/descriptive/**/*.rs` owners such as
`moments.rs` and `reciprocal_sum.rs`, the descriptive test files,
the frozen contract, the numeric oracle/generator and goldens, native
provenance inputs, and the harness. The descriptive owner paths are discovered
from the explicit `descriptive*.rs` and `ods_formula_descriptive_*.rs` globs at
freeze time, so every newly landed candidate Rust or generated contract input
must appear in the final selected-file manifest. Missing owner files fail
closed; no guessed or stale implementation filename is silently substituted.

Capture readiness requires all of the following:

1. The frozen contract hash defines pass counts, signed-mean behavior, domain
   and list-shape behavior, scalar/array/reference result shapes, and
   error/resource precedence.
2. `numeric_oracle.py`, generated numeric goldens, and the contract hash agree;
   the harness has finite expected values or typed failures for every active
   row.
3. Evaluation, limits, native, and oracle tests exist under the descriptive
   test names and pass in the isolated gate checkout.
4. The candidate freeze hashes every selected Rust, test, contract, generator,
   golden, native, harness, and lock input, including newly added owner paths.
5. Root authorizes a quiet window. Only then may the runner build with
   `--locked --offline`, run its mandatory preflight, and take the two-phase
   3-warmup/15-sample capture. No source or profile input may be edited after
   authorization.

The harness now lists and preflights every named row, including typed
cancellation/resource failures and admitted versus refused reference lists.
The runner performs those untimed preflights before any sample. It refuses a
missing contract/oracle input or contract-hash mismatch and records source,
lock, toolchain, preflight, raw stdout, `/usr/bin/time -v`, and cleanup
receipts for both baseline and candidate lanes.
