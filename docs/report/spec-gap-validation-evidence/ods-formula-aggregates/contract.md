# ODF numeric aggregate contract

This is the normative implementation and review contract for the seven numeric
aggregate functions in OpenFormula 1.4 Part 4 §6.16:
`SUM`, `PRODUCT`, `SUMSQ`, `SUMPRODUCT`, `SUMX2MY2`, `SUMX2PY2`, and
`SUMXMY2`. It describes the repository profile where Part 4 leaves a host or
conversion choice open. It applies to the scalar evaluator and the
resolver-backed value evaluator, including inline arrays and supported cell
references. It does not authorize recalculation of formula cells or use of a
cached formula result.

The source is the local distribution:

- archive: `3rdparty/specs/OpenDocument-v1.4-os.zip`
- archive SHA-256: `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`
- member: `part4-formula/OpenDocument-v1.4-os-part4-formula.html`
- member SHA-256: `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`
- relevant sections: §§4.10–4.11.12, 6.1–6.3, and 6.16.47,
  6.16.61, 6.16.64–6.16.68.

The archive hash above is the recorded local source hash. The member hash is
the hash of the uncompressed HTML member and is included so a future review
can verify that the cited text came from the same specification snapshot.

## Signatures and mathematical operations

The signatures below use the specification's parameter notation. `+` means
one or more parameters in the common function template (§6.2); semicolon is
the formula argument separator.

| Function | Part 4 signature | Operation | Shape/arity constraint |
| --- | --- | --- | --- |
| `SUM` | `SUM({ NumberSequenceList N }+)` | Add every selected Number. | Variadic sequence; the §6.16.61 constraint is `N != {}`. |
| `PRODUCT` | `PRODUCT({ NumberSequenceList N }+)` | Multiply every selected Number. | Variadic sequence. |
| `SUMSQ` | `SUMSQ({ NumberSequence N }+)` | Add the square of every selected Number. | Variadic sequence; the §6.16.65 constraint is `N != {}`. |
| `SUMPRODUCT` | `SUMPRODUCT({ ForceArray Array A }+)` | For each matrix position, multiply all corresponding elements and add the products. | Every matrix has the same rows and columns. |
| `SUMX2MY2` | `SUMX2MY2(ForceArray Array A; ForceArray Array B)` | Add `a[m,n]² - b[m,n]²` for every position. | Exactly two equal-shape matrices. |
| `SUMX2PY2` | `SUMX2PY2(ForceArray Array A; ForceArray Array B)` | Add `a[m,n]² + b[m,n]²` for every position. | Exactly two equal-shape matrices. |
| `SUMXMY2` | `SUMXMY2(ForceArray Array A; ForceArray Array B)` | Add `(a[m,n] - b[m,n])²` for every position. | Exactly two equal-shape matrices. |

For `SUMPRODUCT` with `K` matrices, the mathematical result is
`sum[m,n](product[k=1..K](a[k,m,n]))`. A single matrix therefore reduces to
the sum of its elements. No implicit broadcasting, truncation, or
intersection is part of these operations.

## Sequence and array conversion

The distinction between `NumberSequenceList`, `NumberSequence`, and
`ForceArray Array` is observable and must remain visible in both evaluators.

### Number sequences (`SUM`, `PRODUCT`, `SUMSQ`)

For a scalar Number, Logical, or Text argument, the scalar `Number`
conversion in §6.3.5 is used. The repository profile converts Logical
`FALSE`/`TRUE` to `0`/`1`, parses a finite locale-independent decimal Text, and
returns formula `#VALUE!` for malformed or non-finite Text. A scalar Empty
value is converted as the existing scalar Number bridge defines it; an
explicit Missing argument is formula `#VALUE!`.

For a reference consumed as a sequence, §6.3.7 and §6.3.8 apply. The
sequence contains Number cells and Error cells in reference occurrence order.
Referenced Empty and Text cells are omitted. This evaluator distinguishes
Logical from Number, so referenced Logical cells are omitted as §6.3.7/8
requires when the types are distinguishable. A formula Error remains an Error
element and is propagated. A `ReferenceList` is flattened in list order, with
each area traversed row-major; this is the only flattening performed by these
sequence functions.

An inline general Array is already a value rather than a cell Reference
(§§3.3 and 4.10), so its elements are visited in row-major order and use the
element conversion profile. For `SUM`, `PRODUCT`, and `SUMSQ`, an inline
Array element is handled as follows: finite Number is included; Logical is
converted to `0`/`1`; Empty is converted to `0`; Text is parsed by the scalar
finite Text parser; malformed or non-finite Text is generated `#VALUE!`;
Missing is generated `#VALUE!`; and a formula Error is propagated. Complex
elements are generated `#VALUE!` under the existing scalar Number bridge.
This array-element rule is separate from the referenced-cell omission rule
above, which is why a Text or Logical in a referenced `SUM` range is omitted
while the same value in an inline Array is converted.

`SUM` and `PRODUCT` use `NumberSequenceList`, so an ordered list of rectangular
areas is a supported input. `SUMSQ` uses `NumberSequence`; one rectangular
Reference is supported, including a 3-D cuboid represented as one Reference
whose sheet planes are traversed in the specification's sheet order. An
explicit multi-area ReferenceList is a pseudotype mismatch for `SUMSQ` and
produces formula `#VALUE!`. A reference is consumed as a sequence even when it
contains one cell; scalar implied intersection is not applied to these three
functions.

### ForceArray matrices (`SUMPRODUCT` and `SUMX*`)

Every `ForceArray` argument is evaluated in non-scalar array mode, including
when the surrounding evaluator is in scalar mode. A scalar Number, Logical,
Text, or Empty is a one-by-one matrix after the element conversion below. An
inline array keeps its rectangular row and column dimensions. One rectangular
reference keeps its full dimensions and is traversed row-major. A
multi-area ReferenceList is not an Array for this profile and produces formula
`#VALUE!`; it is never silently concatenated or intersected.

Matrix element conversion is deliberately different from reference sequence
filtering:

- Number is used unchanged when finite.
- Logical converts to `0` or `1`.
- Empty converts to `0`, including an Empty cell in a referenced matrix.
- Text uses the same finite scalar Text-to-Number parser. A malformed or
  non-finite Text produces formula `#VALUE!`.
- An explicit Missing element produces formula `#VALUE!`.
- A formula Error is retained and propagated.
- An unsupported cell kind, formula cell, or provider failure remains the
  corresponding typed evaluator failure; it is not replaced by an Empty or a
  cached value.

This is a repository conversion profile because Part 4 specifies the Array
shape and operation but does not prescribe an element conversion for every
Array member. It keeps the §6.3.5 Logical and empty-reference rules available
for matrix inputs while retaining the sequence omission rules for the
NumberSequence functions.

All matrix dimensions are checked with overflow-safe arithmetic. Equal shape
means equal row count and equal column count, including the one-by-one case;
there is no scalar-to-matrix broadcast. A shape violation is formula
`#VALUE!` after the function's arguments have been evaluated. Pair functions
with fewer or more than two arguments also return formula `#VALUE!`.

## Empty arguments and identities

The strict grammar uses `+`, and the two sequence sections state the
non-empty-sequence constraint. SUM and SUMSQ explicitly permit evaluators to
evaluate calls that fail that constraint; this repository records the selected
identity behavior, while PRODUCT keeps strict arity:

- `SUM()` evaluates to additive identity `0`.
- `PRODUCT()` has zero supplied parameters and returns formula `#VALUE!` under
  the strict `+` signature. `PRODUCT(A1:A2)` where the supplied range contains
  only Empty, Text, or distinguished Logical cells is the distinct
  empty-sequence case and evaluates to multiplicative identity `1`.
- `SUMSQ()` evaluates to additive identity `0`.
- `SUMPRODUCT()` has no matrix argument and returns formula `#VALUE!`.
- The three fixed-arity `SUMX*` calls with zero or one argument return formula
  `#VALUE!`.
- An explicit missing argument, such as `SUM(;)`, is a supplied argument and
  returns formula `#VALUE!`; it is not the zero-argument identity case.
- A supplied reference that yields no selected Number cells uses the
  operation identity: `SUM`/`SUMSQ` return `0` and `PRODUCT` returns `1`.
  Empty and omitted reference cells do not become a product zero under the
  NumberSequence rules.

`SUM()` and `SUMSQ()` use their additive identities as the repository's
explicit evaluation of the §6.16.61/65 permission that evaluators may evaluate
expressions failing the non-empty sequence constraint. `PRODUCT()` has no
corresponding permission in §6.16.47 and therefore keeps the common strict
`+`-arity error. The native fixture
`3rdparty/libreoffice-core/sc/qa/unit/data/functions/mathematical/fods/product.fods`
contains a cached `PRODUCT()` value of `0`; that cache is compatibility
evidence only and is excluded from the normative choice of `#VALUE!`.

## Formula-error order and typed failures

The common §6.1 rule applies: an Error supplied by an argument or encountered
in an admitted reference is a formula result, and the leftmost Error should be
returned when several are present. The deterministic repository traversal is:

1. Function arguments are in source order.
2. Sequence areas are in ReferenceList occurrence order, and cells within an
   area are row-major.
3. Matrix arguments use their argument order and row-major cell order for
   conceptual error precedence. The implementation may materialize or read
   matrices in another order, but it must return the same first formula Error.

Conversion, arithmetic, and shape failures are generated formula errors. A
formula Error has precedence over a generated error discovered later in the
conceptual traversal; the evaluator continues the bounded scan when necessary
to find that Error. If no formula Error exists, the first generated error in
the selected traversal is returned. A typed `Unsupported`, `ResourceLimit`,
`Allocation`, `Cancelled`, `SourceChanged`, or `SourceVersionAvailabilityChanged`
failure aborts evaluation and remains catchable only by the host API, not by
`IFERROR`/`IFNA`.

Direct argument Errors are observed before shape or arity constraint results
for an otherwise syntactically valid call. Syntactic parsing failures and an
unavailable required argument are formula `#VALUE!`. A later argument must
not hide an earlier formula Error merely because its matrix has a different
shape.

## Numeric profile

Part 4 gives the mathematical result and permits an implementation to choose
an algorithm that handles finite precision and range (§6.2 and the general
numeric model). The bounded profile has the following observable requirements:

- `SUM` uses a fixed, checked binary accumulator with round-to-nearest-even
  conversion to finite binary64. It preserves exact cancellation between
  values of widely different exponents and retains a representable subnormal
  result. The final exact total is rounded once, round-to-nearest-even, and is
  formula `#NUM!` only when that final binary64 conversion is non-finite or
  otherwise unrepresentable; a total just beyond the finite maximum may still
  round to the finite maximum. NaN and infinity never escape.
- `PRODUCT` keeps sign, zero, normalized mantissa, and exponent separately.
  Intermediate overflow and underflow do not decide the result. A zero factor
  produces signed zero according to the parity of negative factors, including
  negative zero. The existing bounded product kernel may round its normalized
  binary64 mantissa at each factor update, so Product comparisons use the
  documented finite-value tolerance; it must still postpone range decisions
  until the final scaled conversion. The final scaled magnitude is rounded to
  binary64 and is `#NUM!` only when that conversion is non-finite or otherwise
  unrepresentable; an unrepresentable tiny magnitude rounds to zero.
- `SUMSQ` accumulates exact finite pair squares in a wider fixed-point state
  before one final binary64 conversion. It must retain tiny squares that become
  representable after addition and must return `#NUM!` only when the final
  rounded binary64 conversion is non-finite or otherwise unrepresentable. A
  square that would overflow an intermediate binary64 multiplication is not by
  itself a failure.
- `SUMX2MY2` accumulates the exact signed difference of pair squares so equal
  large terms can cancel. Computing `(a-b)*(a+b)` or an equivalent scaled
  form is acceptable; direct overflowing `a*a`/`b*b` followed by subtraction
  is not.
- `SUMX2PY2` accumulates the exact non-negative sum of pair squares with the
  same wide range and final overflow rule as `SUMSQ`.
- `SUMXMY2` forms each finite difference with scaling as needed and accumulates
  its square in the same wide state. A large difference whose final rounded
  binary64 conversion is non-finite or otherwise unrepresentable is `#NUM!`;
  a small difference between very large, close
  operands must not be lost to an overflowing or prematurely rounded square.
- `SUMPRODUCT` forms each K-factor product as a sign, normalized mantissa,
  and exponent and adds terms in an extended cancellation-preserving sum.
  K=1 therefore has an additive reduction, using the ForceArray element
  conversion profile (so it is not necessarily the same as `SUM` on a
  referenced Text or Logical cell). For K=2, it uses an exact fixed-width
  pair-product accumulator. For K>2, the fixed exponent-span accumulator
  retains every term that can affect the admitted result; no term
  is discarded by a scale shift. The implementation profile is a fixed signed
  128-limb window of 64-bit limbs: 8192 bits total, with 64 high bits reserved
  for carry, leaving an 8128-bit admissible input window. The window is a
  property of the numeric profile; increasing the caller's Memory budget does
  not enlarge it. Each K>2 product is normalized and its binary64 mantissa is
  rounded at every factor update. The resulting rounded 53-bit significand is
  inserted at a checked bit offset, carries are retained, and a base shift is
  admitted only while the fixed 8128-bit input window remains sufficient; the
  limb reduction is exact for those rounded product terms. An
  implementation must not silently drop low tails during a large scale shift
  while claiming exact cancellation. If the required span exceeds this fixed
  window, the evaluator returns a typed `Resource::Memory` refusal reporting
  the required and profile-limit bytes (the current two-magnitude admission
  limit is 2,032 bytes); caller work and storage budgets remain additional
  admission checks. For K>1, an intermediate product outside
  binary64 is not by itself `#NUM!`: for example, products of equal magnitude
  and opposite sign may cancel to zero. If the final rounded binary64
  conversion is non-finite or otherwise unrepresentable, the result is
  `#NUM!`; if it is exactly zero, the result is positive zero; representable
  subnormal totals remain subnormal.

All exact accumulators are finite and checked. The result of an exact zero
sum is positive zero for the sum families. The product sign rule above is the
only signed-zero requirement. Cross-platform tests should compare finite
nonzero results within the repository's documented binary64 tolerance; tests
for zero, formula-error kind, subnormal preservation, and signed Product zero
are exact.

## Work, storage, and resolver boundaries

The aggregate evaluator is a scalar-producing, read-only calculation. It must:

- stream NumberSequence and ReferenceList inputs cell by cell, retaining only
  fixed aggregate state and bounded argument metadata;
- materialize only the bounded rectangular matrices required by the current
  value VM, or process paired references in lockstep when the resolver can
  provide stable dimensions and cells;
- check every row, column, area, cell count, and byte product before
  allocation, using `max_reference_cells`, `max_reference_areas`,
  `max_array_cells`, and the storage budget;
- charge AST visits, argument conversion, every physically inspected cell,
  matrix operation, and retained owned text to the caller's work and storage
  budgets; and
- check cancellation during long scans and before publishing the result.

The aggregate state is fixed-size or explicitly bounded. It must not collect
an unbounded `Vec` of all numbers, clone borrowed Text unnecessarily, create a
result matrix for a scalar result, or reset a budget for each argument. A
fallible allocation or a limit refusal is reported as a typed evaluator
failure. Formula errors remain values and can be handled by the existing
formula error functions.

The resolver is immutable and source-version checked before and after the
evaluation. Reference reads do not recalculate formulas, refresh external
data, publish caches, or perform I/O. A formula-bearing or otherwise
unsupported cell is refused according to the existing value-evaluator
capability contract.

These boundaries follow `docs/adr/0004-semantic-api-design.md`,
`docs/adr/0005-io-memory-and-performance.md`, and `docs/GOAL.md`: bounded and
fallible public behavior, hierarchical work/storage/cancellation accounting,
stable read-only source access, and no hidden recalculation.

## Required evidence and regression vectors

The implementation review must cover every function through scalar values,
inline arrays where the parser admits them, and supported rectangular
references. It must include:

- scalar Number, Text, Logical, Empty, Missing, and formula Error cases;
- reference Number, Text, Logical, Empty, Error, single-cell, multi-cell,
  ReferenceList, and row-major order cases, with the NumberSequence versus
  ForceArray distinction visible in expected results;
- zero arguments, an empty selected sequence, explicit missing arguments,
  fixed-arity violations, and matrix shape mismatches;
- first formula Error versus later conversion, shape, and numeric failures;
- large cancellation (`SUM(1e16;1;-1e16)`, difference of equal `f64::MAX`
  squares), overflow-cancellation in `SUMPRODUCT`, and products/squares near
  the smallest subnormal; and
- work, storage, reference-cell, cancellation, unsupported-cell, and source
  version refusals with no partial cache publication.

The independent binary64/Fraction observations in
`docs/report/spec-gap-validation-evidence/ods-formula-aggregates/numeric-goldens.json`
are supplemental numeric evidence. Native LibreOffice FODS files are
provenance and interoperability observations; their cached values do not
override this contract, especially for the zero-argument `PRODUCT()` row.
