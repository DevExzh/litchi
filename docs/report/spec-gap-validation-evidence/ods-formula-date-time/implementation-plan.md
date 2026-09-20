# Date/time implementation and validation plan

Status: preparation; production implementation and validation are pending.
Scope is all 24 functions listed in `contract.md`, not a subset of the family.

## Baseline and isolation

The production baseline is commit `6fa3b8af6a`; the inspected checkout HEAD
`600e86f4e009ebaf57b7245209cedef6f3326e6a` adds lookup evidence only.
The initial reviewed draft is SHA-256
`08e3ce3c3ef34145b2d5963fd2581c63db1c7729f36c9b03a4a62842ab5b40bb`.
This is not a contract freeze: semantic and resource reviewers are reviewing it.

Cargo gates will use a separate checkout and the retained lookup gate lock at
`../ods-formula-lookups/gates/Cargo.lock`, SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The ambient root lock is different, SHA-256
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`,
and must remain untouched. Root coordinates Cargo commands and isolated build
targets. Existing unrelated work and temporary directories are outside this batch.

## Inspected integration boundaries

All paths below are relative to `crates/litchi-ods/src/codec/formula/`.

| Owner | Existing boundary | Required change |
| --- | --- | --- |
| Calendar and parser | `evaluation/calendar.rs`, `evaluation/inspection/parse_value.rs` | Share civil-date conversion and bounded borrowed parsing; retain VALUE's existing date domain while admitting the explicit date-family domain |
| Public options | `evaluation.rs` | Validated immutable calculation timestamp, option builder/accessor, constructor defaults, typed absent-clock capability |
| Scalar kernel | New focused `evaluation/date_time` modules | One function catalog and shared conversion/calendar algorithms for all 24 functions |
| Scalar dispatch | `evaluation.rs`, `evaluation/value/scalar.rs` | Exact arity and missing-slot handling; route both evaluators to shared kernels |
| Value VM | `evaluation/value.rs`, new focused child module | Project ordinary scalar arguments; retain complete holiday/workweek descriptors; stream reads and charge normalized holiday storage |
| Demand cache | Value VM scheduling, shape, reference-kind and criterion classifiers | Propagate full-argument context only to sequence slots; keep nested MUNIT scalar arguments position-sensitive; timestamp belongs to evaluation identity |

The current calendar module has civil-from-days conversion and English month
tables. The existing VALUE parser already receives borrowed text and has
caller-owned scratch for long grouped values. EvaluationOptions is Copy/Eq
with private fields; EvaluationContext::new constructs it explicitly and must
be updated with any new field. UnsupportedKind currently has no clock variant.

## Work ownership

Contract owner and independent semantic/resource reviewers settle normative
and selected-profile questions before production coding. Separate preparation
owners cover calendar/parser/API, value scheduling, test coverage, independent
oracle vectors, and performance scenarios. Production file ownership will be
assigned explicitly to prevent concurrent scalar/value API changes.

The scalar and value owners must agree on shared signatures before removing
helpers or changing callers. A passing package check alone is insufficient:
test-target compilation also enforces dead-code and API consistency.

## Validation and publication

1. Bind the accepted contract to its exact hash and retain independent reviews.
2. Implement the complete family with focused semantic, matrix, resource and
   independently derived oracle tests. Test VALUE compatibility separately.
3. Exercise timestamp presence/absence, eager sequence scans after formula
   errors, typed failure precedence, zero-read shape refusals, cancellation,
   source fences, allocation bounds and reservation drop order.
4. Freeze the selected production, test, contract and oracle files; retain
   their hashes, isolated lock and reproducible gate commands.
5. Run isolated package tests, strict Clippy/rustdoc, formatting, crate
   boundaries and diff checks. Record actual counts and terminal results.
6. Run native corroboration and matched baseline/candidate performance using
   a fixed harness after semantic/read-count preflight. Retain raw samples,
   allocation/work/read metrics and any measured regressions; do not claim
   workbook speedups from evaluator microbenchmarks.
7. Independently verify retained evidence, commit the completed production
   batch and supporting evidence, then remove only owned disposable targets,
   checkouts and caches. Retain diagnostics needed to explain evidence.

No production-support, test-PASS or performance claim follows from this plan.
