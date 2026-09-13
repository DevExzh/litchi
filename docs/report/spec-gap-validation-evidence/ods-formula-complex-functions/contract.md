# Complex-function contract

This contract is the normative input for a future implementation of the ODF
1.4 Part 4 complex-number family. The source is the local
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` member of
`3rdparty/specs/OpenDocument-v1.4-os.zip` (archive SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`; HTML SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`).

## Scope and function set

The complete §6.8 family is:

`COMPLEX`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`,
`IMCOSH`, `IMCOT`, `IMCSC`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`,
`IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSIN`, `IMSINH`, `IMSEC`,
`IMSECH`, `IMSQRT`, `IMSUB`, `IMSUM`, and `IMTAN` (§6.8.2–§6.8.27).

The formula catalog already recognizes these names, but the current scalar
and value evaluators dispatch unsupported function names as typed
`Unsupported(Function)`. A complete implementation must cover the family in
each evaluation profile that advertises it. It must not add names to the
dispatch table while leaving value, array/reference, or owned-result paths
unable to carry the result.

## Representation and conversion

Part 4 §4.4 defines a complex number as a pair of real and imaginary parts and
allows it to be represented as Number, Text, or a distinguishable type. The
recommended internal representation is a fixed-size pair of finite `f64`
components. If complex values are observable through the public evaluator,
add a distinguished public value variant and update borrowed views, owned
conversion, arrays, structural equality, and result storage together. Text
round-tripping should not be used as the internal representation: it adds
locale, formatting, allocation, and loss-of-sign/branch risks.

Conversion follows §6.3.10. A non-complex Number becomes that real value with
zero imaginary part; a Logical converts through Number; a Text is converted
using the complex forms accepted by `VALUE`, including the two forms
`([+-]? Number [+-])? Number [ij]` and `[+-]? Number [ij]`; a reference is
converted to Scalar first and an empty reference is zero. `COMPLEX` itself
accepts Real, Imaginary, and an optional suffix; the suffix policy in §6.8.2
is lowercase `i` or `j`, and invalid suffixes remain formula errors.

For `ComplexSequence`, §6.3.11 makes a scalar Number, Text, or Logical a
one-element sequence. A reference contributes Number, Text, and Error cells
in the specified reference-list/cuboid order; Empty cells are omitted. Since
this evaluator distinguishes Logical from Number, logical cells in a
reference sequence should be omitted unless a later profile explicitly
chooses the non-distinguished interpretation. Errors must retain formula
error precedence rather than being silently converted or dropped.

## Function-specific choices that must be explicit

`IMSUM` has two special rules: unconvertible Text is ignored, and zero
parameters are implementation-defined (Error or Number 0). Choose and test
one policy. Do not generalize that ignored-Text rule to other complex
functions. `IMARGUMENT(0)` is also implementation-defined (0 or an Error).

The local §6.8.24 `IMSQRT` text is internally problematic. It declares a
Complex result, but its displayed equation places `sin(argument/2)` in the
real component and `cos(argument/2)` in the imaginary component. The polar
definition in §4.4 and the usual principal square root place those terms in
the opposite components: `sqrt(r) * (cos(phi/2) + i sin(phi/2))`. For
example, the literal equation makes `IMSQRT(1)` equal to `i`, while the
principal square root is `1`. This must be an explicit implementation-policy
decision with regression vectors; implementation must not claim exact
conformance without recording which reading is used.

The local §6.8.23 `IMSECH` entry declares `Returns: Number` but defines the
result as `IMDIV(1; IMCOSH(N))`, whose complex division result is generally
Complex. The public result type and conversion policy likewise require an
explicit decision. A standard complex-valued result follows the displayed
formula; a Number result would need a documented projection rule and would
discard information.

Use deterministic `atan2` argument range `-π < phi ≤ π`, and document branch
and signed-zero behavior for logarithm, square root, and power. Enforce the
specified nonzero domains for `IMLN`, `IMLOG10`, `IMLOG2`, and `IMPOWER`, and
denominator rules for `IMDIV`. Use scaling where a finite result is representable despite overflow or
underflow in a naive intermediate. Non-finite final components should map to the evaluator's formula-level numeric error (`#NUM!`) rather
than escaping as NaN/Infinity. Formula errors remain values; cancellation,
source/provider failures, and resource exhaustion remain typed evaluation
failures.

## Bounded evaluation and tests

`IMSUM` and `IMPRODUCT` should stream their `ComplexSequence` inputs rather
than collect an unbounded temporary vector. Charge work per converted element,
bound reference cells, text parsing, stack, and owned storage, and check the
retained `ExecutionContext` during long sequences and before fallible
allocation. A scalar complex operation should require no heap allocation.
Complex text parsing must obey the existing text-byte limit, and malformed
or over-limit input must be refused before proportional allocation.

Tests should cover every function above, both `i` and `j` forms, real-only and
imaginary-only values, invalid text/suffix/arity, zero and domain boundaries,
overflow/non-finite results, error ordering, reference-list and cuboid order,
Empty/logical sequence policy, `IMSUM` ignored text, zero-arity choice,
`IMARGUMENT(0)`, and the selected `IMSQRT`/`IMSECH` interpretations. Include
work, storage, text, and cancellation refusals and verify that owned results
retain no input lifetime.

Database functions remain a separate later family: `DAVERAGE`, `DCOUNT`,
`DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DSUM`,
`DVAR`, and `DVARP` (§6.9.2–§6.9.13) require a database/criteria execution
layer under §§4.11.8–4.11.11, rather than only numeric scalar kernels.

## Selected implementation policy

The bounded profile uses the mathematical principal square root for `IMSQRT`,
consistent with §4.4 and the function summary, and a Complex result for
`IMSECH`, consistent with its defining division. These are explicit resolutions
of the contradictory local specification text, not claims that its conflicting
sentences can both be satisfied. Regression vectors must include `IMSQRT(1)`,
`IMSQRT(-1)`, and non-real `IMSECH`.

`IMARGUMENT(0)` returns Number 0. `IMSUM()` returns Number 0; `IMPRODUCT()`
returns an arity error. Nonzero logarithm/power constraints are enforced,
including `IMPOWER(0;0)` returning a formula error. Complex reference sequences
omit Logical and Empty; scalar Logical converts through 0/1. No host locale,
ambient provider, or external execution is introduced.

## Independent extreme-value oracle

Root recomputed these constants using Python `decimal` with precision 80,
using `r = sqrt(a*a+a*a)`, `ln(r)`, and square-root components
`sqrt((r+a)/2)` / `sqrt((r-a)/2)` for decimal `a = 1e308`:

- `ln(abs(1e308+1e308i))`: `709.5427822324460433322499841035`.
- Principal square root real: `1.0986841134678099660398011952e154`.
- Principal square root imaginary: `4.5508986056222734130435775782e153`.
- `abs(1.2e308+1.2e308i)`: `1.6970562748477140585620264691e308`.

These are decimal mathematical inputs; binary-f64 tests should allow an
appropriate small relative rounding tolerance. A reviewer draft gave an
incorrect logarithm approximation and repeated the square-root real component
as its imaginary component; those draft numbers are not test oracles.
Division of either `1e308+1e308i` or `1e-308+1e-308i` by itself must remain
approximately `1+0i`, without the naive square-sum overflow/underflow.
