# Candidate 09 review

Root integrated independent coder, tester, numerical, boundary, and profiler
reviews before freezing the 464-source candidate. This is a scoped review of
the complex family; it does not establish completion of the full spec audit.

Numerical review led to scaled division, product, logarithm, square-root,
exponential/hyperbolic, and trigonometric reciprocal kernels. Root added
independent extreme-result regressions and corrected a negative-zero product
sign found in the last numerical review. All 30 focused tests pass.

Boundary review checked Complex propagation through cached/projected values,
matrix transposition, ordinary coercions, streamed reference aggregation, and
owned copies. A scalar argument buffer now carries its reservation through
its consuming iterator. Scalar aggregates retain the first generated error,
pre-scan original formula errors, check cancellation periodically while popping
and scanning, and charge/check each fold element. Tests verify scalar/value
error ordering and typed Work refusal. No new public archive, resolver, or
expression lifetime leaks into owned Complex values.

The profiler independently verified the release build maps against candidate
source maps, all 675 raw capture-artifact hashes, 225 successful runs, and the
unchanged retained binary. Result checksums, allocation/request/release/peak
counters, work, and resolver counts are deterministic across its three rounds.
All four typed refusal cases match their expected categories. Retained result
budget memory is zero in this scalar-result corpus; allocator activity and
process RSS remain distinct measurements.

The existing-value comparison has zero deterministic-field mismatches but
reports six RSS flags and one +16-byte requested/released sample delta. These
are retained in the performance report for follow-up; this review does not
convert diagnostic measurements into broad performance acceptance.

Root verified current canonical source hashes still equal the passing
candidate after the performance captures. All five ODS gates pass. Raw logs
retain their exact original trailing blank lines; source/document diff checks
pass without rewriting captured evidence. Temporary capture/build receipts
were archived and byte-verified before loose scratch directories were removed.
