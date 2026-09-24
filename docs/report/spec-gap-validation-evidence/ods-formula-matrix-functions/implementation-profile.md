# Matrix evaluation profile

This records implementation choices for ODF 1.4 Part 4 §6.5 where the
[normative contract](contract.md) leaves behavior unspecified. It is an
implementation target; it does not claim that the functions or their
validation gates are complete.

`MDETERM`, `MINVERSE`, and `MMULT` accept finite Number elements. Text,
Logical, Empty, and missing elements produce a formula `Value` error.
Existing formula errors propagate in argument order and then row-major
element order. This element policy is distinct from converting a scalar
Number parameter: the specification declares these parameters as Array,
without defining a numeric element conversion.

`TRANSPOSE` preserves element types and exchanges rows and columns. Its result
has no reference origin: consuming that result uses array semantics, not the
implicit intersection rules of its input reference. Error elements retain
their positions after transposition; a scalar Error argument propagates.

`MUNIT` uses Number parameter conversion followed by truncation toward zero.
Logical values convert to zero or one, Empty converts to zero, and Text
produces a formula `Value` error, including numeric-looking text. The
resulting integer must be positive. Array-valued input to its scalar
parameter supplies the first element rather than triggering repeated calls
that would create an array of matrices (§3.3, rule 2.2.1). Non-finite numeric
results produce a formula `Number` error. Checked dimensions and cell-count
admission happen before output allocation; exceeding a host limit produces
an operational resource failure, not a catchable formula error.

Determinant and inverse use elimination with partial pivoting, with cubic
arithmetic work and quadratic temporary storage. A zero pivot identifies a
singular matrix: determinant returns zero and inverse returns a formula
`Number` error. The implementation does not add an arbitrary singularity
tolerance. Finite floating-point roundoff is permitted; non-finite arithmetic
results produce a formula `Number` error. This does not promise reliable
inversion of ill-conditioned matrices beyond the evaluator's binary64
numerical model.

The evaluator charges arithmetic work inside bounded loops, checks
cancellation, and retains storage reservations for the lifetime of their
buffers. Unselected lazy branches do not evaluate matrix arguments or run
their numerical kernels. Matrix support remains within the existing
read-only, source-version-checked value API; it does not activate recursive
worksheet recalculation or trust formula caches.

Nested value probes used by shape discovery have a structural ceiling of 32
active probes, further bounded by `max_stack_entries` and the caller's Depth
budget. Exceeding it returns an operational `ResourceLimit(Depth)` and
releases the suspended storage. Ordinary expression traversal remains
iterative. This ceiling bounds the remaining probe call-stack use; it is
distinct from the mathematical dimensions of a matrix.
