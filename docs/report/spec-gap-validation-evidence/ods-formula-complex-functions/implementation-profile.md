# Complex evaluation profile

Both evaluator entry points implement the 26 OpenFormula §6.8 functions listed
in [the contract](contract.md). This is one function family within the bounded
constant/reference evaluator, not dependency recalculation or complete ODF
Small Group support.

## Representation and ordinary operators

`evaluation::complex::Complex` is a fixed-size `Copy` value containing finite
real and imaginary components and an `i`/`j` formatting suffix. Its checked
constructor returns a typed error for non-finite components or an invalid
suffix. Scalar, value, array, and owned results expose a Complex variant.
Equality compares the components and ignores the suffix. Different scalar
kinds compare unequal; ordinary ordered, numeric, and logical coercions refuse
Complex with a formula Value error, including real-only Complex values.
Unary plus preserves Complex and unary minus negates its components.
Concatenation formats it through bounded stack scratch and an explicitly
reserved output string. A real-only textual result does not include a suffix.

The branch range is `(-π, π]`; the negative real axis uses `+π` even for a
negative-zero imaginary component. Square root follows the mathematical
principal branch. The selected zero, sequence, and contradictory specification
text policies are recorded in the contract. No locale or external provider
is implicitly consulted.

## Numerical and resource behavior

Multiplication and division retain component products as binary mantissa and
exponent terms until final scaling. Logarithm avoids constructing an overflowing
modulus. Square root derives its smaller component from the larger one to
avoid cancellation. Large trigonometric reciprocals use exponential ratios;
exponential and hyperbolic components scale their coefficients before the
final exponentiation. Non-finite final components are formula Number errors.
The tests include finite extreme results, tiny components beside huge ones,
principal branches, and signed zero. This is binary-f64 arithmetic, not an
arbitrary-precision or correctly-rounded transcendental contract.

Complex kernels allocate no component storage. Existing evaluator stacks and
results remain budgeted. Fixed-arity calls pop their small argument set directly.
Scalar variadic calls retain an admitted argument buffer whose reservation
outlives the buffer and consuming iterator. Resolver-aware sequences stream
references and arrays through a fixed-size accumulator; they do not collect
all referenced cells. Original formula errors precede generated errors, and
the first generated error is retained. IMSUM ignores unconvertible text.
Referenced Empty and Logical cells are omitted; explicit scalar Logical
arguments convert through zero and one. Reference-list duplicates remain
ordered inputs.

Work, cancellation, text, reference, shape, stack, and owned-storage limits
remain explicit. Copying a scalar Complex to owned storage needs zero additional
reserved bytes; owned arrays still reserve their element storage. The copied
value contains no expression or resolver lifetime. Matrix numeric functions
reject Complex, while TRANSPOSE preserves it.

## Validation scope

Exact candidate sources, test commands, and gate outcomes are retained alongside
this profile. The performance harness measures the new family independently;
unsupported pre-implementation calls are not an equal-work latency baseline.
Existing-workload comparisons and outstanding performance flags must be reviewed
separately before making a broad no-regression claim.
