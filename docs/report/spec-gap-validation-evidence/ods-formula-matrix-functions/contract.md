# Matrix-function contract

This contract is the normative input for implementing ODF 1.4 Part 4 §6.5. The source is the local `part4-formula/OpenDocument-v1.4-os-part4-formula.html` member of `3rdparty/specs/OpenDocument-v1.4-os.zip`.

## Common array semantics

A matrix has `M` rows and `N` columns; a general array has equal-width rows and may have no value at a row/column intersection (§6.5.1, §4.10). Inline arrays use `{ ... }`, have at least one row and one column, and require the same column count in every row. Their elements are constant expressions and the result has type `Array` (§5.13). The specification warns that inline elements other than constant Number or String can reduce interoperability; it does not make those other element types impossible.

`ForceArray` is an argument attribute, not a value type. It evaluates that argument in non-scalar array mode: no implied intersection is performed, and a multi-cell reference is iterated into an array (§6.3.4). Without `ForceArray`, non-scalar arguments follow §3.3: scalar use takes the `[0,0]` inline-array element or performs reference implied intersection; matrix mode iterates scalar-argument functions with rectangular broadcasting and out-of-range `#N/A`. For an array-producing function whose scalar-typed argument is an array, matrix mode does not implicitly iterate that call over the argument; it uses the argument's `[0,0]` element to make one call (§3.3 rule 2.2.1, illustrated by Note 7). Thus `MUNIT({2;3})` uses `N=2` and returns one 2 × 2 array. This input rule does not collapse an array returned for later element-wise use: `TRANSPOSE(...)+1` still maps the full transposed result. Matrix implementations must preserve these context boundaries.

The signatures are:

| Function | Signature | Result and required shape rule |
| --- | --- | --- |
| `MDETERM` (§6.5.2) | `MDETERM(ForceArray Array A)` | Number, the determinant; `A` shall be square. |
| `MINVERSE` (§6.5.3) | `MINVERSE(ForceArray Array A)` | Array, the inverse; `A` shall be square. A singular matrix should produce an Error. |
| `MMULT` (§6.5.4) | `MMULT(ForceArray Array A; ForceArray Array B)` | Array of shape `rows(A) × columns(B)`; `columns(A) = rows(B)` and each output is the sum of element products. |
| `MUNIT` (§6.5.5) | `MUNIT(Integer N)` | `N × N` identity Array; `N` shall be greater than zero. |
| `TRANSPOSE` (§6.5.6) | `TRANSPOSE(Array A)` | Array with exchanged dimensions (`N × M`); no additional constraint. |

## Element conversion, missing values, and errors

The matrix sections define the mathematical operations but do **not** declare a per-element `Number` pseudotype or specify how Text, Logical, Empty, or missing Array elements become matrix numbers. §6.2 generally requires conversion to the expected parameter type, but `Array` does not itself define an element conversion. Therefore a conforming implementation must make this profile choice explicit rather than silently claiming that §6.5 mandates it.

If the implementation applies §6.3.5 element-wise for numeric matrix operations, Number is unchanged, Logical converts to `0`/`1`, and Text conversion remains implementation-defined and locale-sensitive. The blank-reference rule in §6.3.5 (`0`) does not resolve an absent intersection in a general Array (§4.10). Empty/missing elements therefore need an explicit policy (for example, a formula Error); they must not be assumed to be zero without documenting that extension. Unsupported element types likewise need a documented Error policy.

An Error supplied as an argument normally propagates, and multiple Errors should return the leftmost one (§6.1; §3.2.3). None of the five matrix functions suppresses this rule. “Leftmost” is not defined for a two-dimensional element traversal, so the implementation should choose and document a deterministic order (row-major is the natural profile). Dimension violations are Errors under the common function template (§6.2–§6.3.1). A singular `MINVERSE` result is worded “should return an Error,” rather than “shall,” in §6.5.3; this is an explicit conformance choice.

`MUNIT` receives `Integer`; after conversion to Number, conversion of a non-integer is implementation-defined when the function supplies no rounding rule (§6.3.6). §6.5.5 supplies no such rule, so the profile must choose (reject, truncate, floor, etc.) before implementation. Zero and negative values fail the `N > 0` constraint. Numerical precision, overflow, and algorithm choice follow the numerical-model latitude in §3.6 and the common template note in §6.2; they must be distinguished from typed resource-limit refusals imposed by the host evaluator.

ODF gives no matrix-dimension or element-count maximum in §6.5. The implementation may impose finite work, stack, storage, and array-cell limits, but must check shape/size before allocation and report those host limits separately from formula Errors. No rule in §6.5 authorizes formula-cell cache trust, recursive recalculation, or external I/O.

## Recommended initial profile choices

For a deterministic first implementation, require finite Number elements for the five matrix operations, propagate Error elements, and return a formula Value error for Text, Logical, Empty, or absent elements. This avoids silently choosing locale-sensitive Text conversion and avoids extending the blank-reference-to-zero rule to general Arrays. A broader Number/Logical/Empty profile can be added later only with explicit tests and documentation; §6.3.5 permits Logical `0`/`1` and leaves Text conversion implementation-defined, but §6.5 does not require either policy.

Use truncation toward zero for `MUNIT`'s unspecified Number-to-Integer conversion, matching the existing evaluator profile, then enforce `N > 0`; reject nonfinite values rather than saturating. Keep dimension violations, singular inverse, and invalid elements as formula Errors, while preserving typed host resource-limit and cancellation failures as distinct outcomes.
