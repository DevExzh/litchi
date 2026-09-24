# ODS conditional aggregate validation

This batch implements `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`,
and `AVERAGEIFS` in the reference-aware value evaluator. These functions
consume cell references and compile criteria once before streaming the selected
cells. The [contract](contract.md) records reference admission, geometry,
criteria conversion, host matching properties, error ordering, and resource
behavior. It also explains the resolver-free evaluator's reference boundary.

The [independent numerical oracle](numeric_oracle.py) uses Python's standard
library `Fraction` over the exact represented binary64 cell values. Selection
masks are computed independently from integer criterion columns; the oracle
never reads evaluator output. Its [288 observations](numeric-goldens.json)
cover all six functions, with 48 input sets including 40 seeded sets. The Rust
integration test checks every observation in scalar and matrix modes of the
reference-aware evaluator, requiring exact final binary64 bits. Cases include
finite cancellation at the largest exponent, subnormal and signed-underflow
averages, overflowing sums with finite averages, no matches, and destination
ranges expanded from one-cell anchors.

The semantic and resource tests separately cover criteria types, reference
geometry, errors, lazy selected reads, cancellation, budget refusal, and
result lifetimes. Formula caches remain inert; this batch does not add
recalculation or result publication.

The [native evidence](native/README.md) retains 64 numeric observations across
all six functions from eight pinned LibreOffice inputs. Five function fixtures
supply existing FODS caches; `SUMIFS` uses the value emitted by the recorded
LibreOffice conversion of the pinned XLS fixture. The evidence distinguishes
that conversion from preservation of the original XLS cache. Profile
exclusions record host matching differences and unsupported source closures.
The [independent reproduction receipt](native-reproduction.json) records raw
input hash checks, exact regenerated JSON, the converter identity, and
successful temporary-tree cleanup.

Sharing the matcher also corrects bare numeric-looking Text criteria in the
existing database functions. Under §4.11.8, bare `"3"` remains Text and matches
a Text cell containing `3`; `"=3"` parses the operator's operand as Number and
matches numeric `3`. The new regression evidence covers both families. This is
an intentional semantic correction, separate from the unchanged host choices
for case sensitivity, whole-cell matching, regex, and wildcards.

To verify the retained receipts from the repository root, run
`python3 docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/verify.py`.
The verifier regenerates the numerical oracle, checks the native reproduction
receipt, verifies the seven gate logs and frozen source hashes, and recomputes
performance medians from the raw samples. It requires the batch's source files
to match their recorded hashes. The native `reproduce.py` helper separately
refetches the pinned inputs and requires the recorded LibreOffice converter.

The [independent semantic review](review.md) and [resource review](resource-review.md)
cover the frozen implementation. The resource fixes charge range geometry,
array-error inspection, and destination lookup attempts before bounded work;
metadata and criterion text retain their existing ownership limits. Repeated
sheet metadata lookups for selected destinations remain a bounded optimization
opportunity and are included in the measured workloads.

Final isolated validation passed all seven checks with unchanged compiled
sources: 1,313 ODS tests (zero failed or ignored), all-target Clippy with
warnings denied, rustdoc with warnings denied, crate formatting, selected-file
formatting, crate boundaries, and diff whitespace checks. The [gate receipts](gates/)
retain the exact dependency lock, source closure, selected-source freeze,
commands, output logs, and independent hash checks.

The final [performance report](performance/results/performance-report.md) retains
270 baseline and 990 candidate samples: 15 fresh processes per case and phase,
with three warmups. Independent verification confirms unchanged matched
allocation counts, requested bytes, and peak live bytes. Unrounded raw median
time shifts are −3.17% to +4.30%; process RSS shifts are −2.56% to +4.51%.
Nested conditional resolver reads are 704, 2,816, and 11,264 at the three
fixture sizes, consistent with a linear scan. These are bounded evaluator
measurements, not whole-workbook recalculation claims. See the
[root verification](root-verification.json) and [cleanup receipt](cleanup.json).
