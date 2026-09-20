# ODS value inspection performance profile

This profile compares the value-inspection candidate with baseline commit
`d623f3c2ecc0c837017f700174656f0e443759a5`. The baseline gate lock is the
retained byte-profile copy with SHA256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`; the
root `Cargo.lock` hash is recorded separately and must not replace it.

The candidate matrix has 28 matched controls and 59 candidate cases covering
all sixteen names. Scalar lanes cover raw empty/text/logical/number/error
identity and conversion, while reference lanes exercise borrowed mixed-type
cells in row-major order. Dedicated scalar-result lanes cover `TYPE`'s full
rectangular scan, `N`'s explicit intersection, and `N`'s inline array
`[0,0]` selection. Matrix lanes cover elementwise predicates and conversions.
Other lanes cover `VALUE` ISO date/time and mixed-fraction parsing,
`NUMBERVALUE` separator/percent/grouping transforms, invalid-separator and
reference-list refusal before reads, lazy projected `IF` coordinates, sticky
cancellation, and zero-budget resource failure.

Every admitted reference cell is read through the resolver and counted. The
oracle must check exact typed values, array shape and coordinates, formula
errors, and the exact read bound before timing. A known refusal must check
zero resolver reads. The cancellation lanes use four internal repeats and must
observe one successful read before the cancellation supersedes any retained
formula value. Resource/provider/source failures remain typed evaluator
failures and are never converted by inspection functions.

The process harness records elapsed time, allocator calls and bytes, peak live
bytes, execution work, retained budget, resolver reads, input/output bytes,
checksums, and external RSS. Three warmups and fifteen fresh child samples in
both `evaluate` and `parse-evaluate` are the intended final protocol. Failed
or preliminary captures must be retained under a diagnostic directory rather
than overwritten.

Contract gates before final preflight are:

1. accepted argument pseudotypes and zero-read refusal for every function;
2. formula-error inspection versus propagation, including `N`, `TYPE`, and
   `VALUE`;
3. exact `TYPE` array/reference metadata codes;
4. scalar versus elementwise matrix shape and lazy-`IF` coordinate behavior;
5. `ISEVEN`/`ISODD` numeric conversion and non-finite rules;
6. `NUMBERVALUE` separator, grouping, whitespace, percent, and sign grammar;
7. `VALUE` date/time/datetime, mixed-fraction, currency/grouping, and
   deterministic two-digit-year profile; and
8. `NA` arity and formula-error identity.

The current smoke build and evaluate/parse-evaluate preflight cover all 87
named cases. Final capture may run only after these gates are frozen, the
source closure is staged against the candidate freeze, and the independent
gate review records the final contract hash.
