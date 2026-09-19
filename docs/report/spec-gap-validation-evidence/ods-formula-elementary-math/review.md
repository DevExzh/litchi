# Independent ODF 1.4 elementary-math review

## Disposition

The frozen production source and focused validation both pass for the bounded
binary64 profile. The batch implements `ABS`, `EXP`, `LN`, `LOG`, `LOG10`,
`MOD`, `POWER`, `QUOTIENT`, `SIGN`, `SQRT`, and `SQRTPI` in both the
resolver-free scalar evaluator and the value evaluator's array/reference
bridge. The final isolated gate run passed all seven checks: 1,211 ODS tests
passed, with zero failures and zero ignored tests. Its source manifests are
identical before and after the run.

The bounded performance harness was also reviewed before capture. It contains
the elementary sources and tests in its selected source list, uses 30 matched
control phase-groups and 42 candidate cases (84 candidate phase-groups), and
parses cleanly. The capture and its retained report are separate evidence and
must be verified before making a performance claim.

## Source basis and identity

The normative source is the local ODF 1.4 Part 4 formula document recorded in
`contract.md`:

| Input | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| extracted Part 4 HTML member | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| `contract.md` | `16345fc715390342523ef4171afa45e859cb62e2c8151d15b10475e5b6f9ec7d` |

The exact frozen production and focused-test inputs are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `0c40890eed31c691ee106ced38d188b2955c77fb0701d91a5b1041692598a0f9` |
| `crates/litchi-ods/src/codec/formula/evaluation/elementary.rs` | `ebea50eb7e1ce17a325472c3992e9982638f6d902bb41a44367e1ed07a3f2742` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `734f41734c20f1961a2c46a624f6ff1766250f343e949c7b5b9469dd49c11867` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `20b9849af0be790cd5aef1061dcd669fd5314ffda777f0693dbecbc7fee3d638` |
| `crates/litchi-ods/tests/ods_formula_elementary_evaluation.rs` | `66e6dd88b94feb02af3114ae8a631ecb578d70477d0a4084cc3d51f30070461d` |
| `crates/litchi-ods/tests/ods_formula_elementary_arrays.rs` | `e7f3ebd3524f60e73d069d44ec2b159cde8c6014f3f7713535c4130aeff36243` |
| `crates/litchi-ods/tests/ods_formula_elementary_native.rs` | `dfb0cef5388cc1b58aac960040dbbe5af57be3a47c8b3fe306b2ff6eae253458` |
| elementary native cache | `7c1415ce12910e770208fdb9065f9baf980b5138c19d78d2d33e1e2f74237225` |
| retained gate freeze | `30489d9e210fbc44751dfffc1676bb667b3d51ca7279b59452aca9b2241a3275` |

The freeze names baseline commit `8ef0057e5eca7c136c85c71c64b816bc3b520c37`.
The retained gate source-before and source-after manifests both hash to
`7548b4de94d5dd65fa66e504876e50afc25c72922e660b686e152766ca585c57`.
The gate result and verification receipts hash to
`3670094580d69cf16e6eac7b99ef5d979a8bce81acb7905f5c705072655124ad` and
`f71db36e9081ea10a985084bc07b7356618d0ad89f7fe6bf6b51568c19756dc9`,
respectively.

## Semantic checks

The dispatch enum recognizes exactly the eleven names in this batch, and the
same scalar kernels are reached from scalar calls, matrix broadcasting, and
materialized references. `POWER` and infix `^` share `power_result`, preserving
negative-base parity and the selected `0^0 = 1` profile. Overflow, underflow
that leaves a finite zero, and negative-base fractional powers follow the
documented finite-result and domain rules.

`LOG` accepts one or two arguments. An omitted base defaults to 10, while a
syntactically supplied empty base remains `#VALUE!`; zero, negative, or unit
bases return `#NUM!`. Formula errors are checked before number conversion and
in source order. Required wrong arity is `#VALUE!` after the evaluator has
consumed the supplied arguments.

`MOD` uses the already-rounded binary64 remainder and adjusts a nonzero result
only when its sign differs from the divisor. This handles all four dividend
and divisor sign combinations without quotient/product cancellation or an
always-positive `rem_euclid` result. Exact zero is canonicalized to `+0`.
`QUOTIENT` remains independent and truncates the finite quotient toward zero.
The large-quotient vectors include both divisor signs and the excluded native
producer variance `MOD(26^15;77)` is not used to redefine the profile.

`SQRTPI` factors the operation as `sqrt(N) * sqrt(PI)` before the finite-result
check. This keeps the defined result for a maximum finite input and avoids
unnecessary underflow at the minimum subnormal. `LN`, `LOG`, and `LOG10` use
strict positive-domain checks; `SIGN` uses comparisons so both signed zeros
produce numeric zero. All kernels publish only finite `f64` Number values.

## Evaluation boundary

The value bridge keeps an explicit syntactically missing argument distinct
from an Empty worksheet cell. Empty cells use the existing Number conversion,
while a missing required parameter becomes a formula `#VALUE!`. Formula domain
and division errors remain per-cell values in a complete matrix result.
Cancellation, resource exhaustion, allocation failure, provider failure, and
unsupported capabilities remain typed evaluation failures and cannot publish
a partial array or be caught by formula error handlers. Lazy `IF`, `IFERROR`,
and `IFNA` branches continue to avoid resolving unselected references.

The implementation adds no public API, dependency, resolver, host service, or
heap allocation in the numeric kernels. Scalar references remain typed
unsupported results, while value mode applies implicit intersection, shape
broadcasting, and reference limits. The focused tests cover work, storage,
reference-cell, shape, cancellation, and owned-result behavior.

## Independent validation

The focused final gate log reports:

```text
ods_formula_elementary_evaluation: 10 passed; 0 failed; 0 ignored
ods_formula_elementary_arrays:      7 passed; 0 failed; 0 ignored
ods_formula_elementary_native:      1 passed; 0 failed; 0 ignored
```

The native receipt has 56 selected cached observations covering all eleven
functions. It is corroborating cached data, not a native application
execution, resave, acceptance, or cross-platform bitwise-accuracy claim.
The independent numeric oracle and goldens are hashed
`688cddd99e2558e61fd680912f6360a0fc4eeef379d7b6443cf1f78036e1e063` and
`206fb4d44b9b8700b30ec2774ad93b404f0b2a25a5b07f70f1f7f29537d2d737`.
The goldens cover subnormals, extreme finite values, near-one logarithm
bases, large-quotient remainders and the factored `SQRTPI` boundaries. The
authored integration tests additionally cover strict domains, signed zeros
and negative powers.

The final locked gate set passed ODS tests, strict Clippy, rustdoc with
warnings denied, crate and batch formatting, crate-boundary checks, and diff
hygiene. No production source was changed during this review.

## Performance harness custody review

The harness uses fresh child processes, three warmups, fifteen samples per
case/phase, fixed repeats, an immutable in-memory resolver, and a direct
untimed `f64`/shape preflight. The baseline captures 15 control cases across
two phases (30 groups); the candidate captures 42 named cases across two
phases (84 groups). Allocator calls, requested/released bytes, live-byte
balance, peak live bytes, execution work, retained budget memory, and external
RSS are recorded. The runner hashes the selected elementary production files,
focused tests, native cache, complete Rust/Cargo source closure, harness,
toolchain, and lockfile before and after capture.

I checked the performance Python inputs with the AST parser and verified that
`SOURCE_FILES` includes the three elementary tests and native cache. The
performance verifier is fail-closed on source custody, sample cardinality,
checksums, allocator balance, lock/toolchain identity, and candidate case
coverage. It does not claim whole-workbook recalculation, native producer
acceptance, or a language-level resident-memory bound.

Performance timing and memory conclusions are outside this source/harness
review. The final retained capture is checked separately by
[root verification](root-verification.json); this production review has no
semantic or validation blocker.
