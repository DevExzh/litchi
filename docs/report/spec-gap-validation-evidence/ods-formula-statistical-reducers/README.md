# ODS core statistical reducer validation

This batch covers `COUNT`, `COUNTA`, `COUNTBLANK`, `AVERAGE`, `AVERAGEA`,
`MIN`, `MAX`, `MINA`, and `MAXA`. The [contract](contract.md) records their
distinct sequence, reference, coercion, empty-value, and error rules. Evaluation
is explicit and read-only; formula caches remain inert during package CRUD.

The [independent oracle](numeric_oracle.py) generates [576 observations](numeric-goldens.json)
from 64 typed reference datasets. Eight numerical boundary sets and 40 seeded
sets cover extreme cancellation, finite averages of overflowing intermediate
sums, subnormals, negative underflow, and widely separated exponents. Sixteen
mixed datasets cover Numbers, Logical values, Text (including empty and numeric
Text), Empty cells, and formula Errors. Python `Fraction` computes averages
from the exact represented binary64 operands with one final nearest-even
conversion. Counts and extrema are computed independently from the typed cells.
The generator never reads Rust evaluator output.

The Rust oracle test evaluates every retained observation in both scalar and
matrix modes of the reference-aware evaluator and compares exact result bits
or the expected formula error. Separate semantic and resource suites cover the
resolver-free API, argument admission, lists, arrays, 3D references, projection
caching, cancellation, limits, and retained result ownership.

The [native receipt](native/README.md) corroborates all nine functions with
54 numeric observations from nine pinned LibreOffice FODS inputs. It uses
existing caches with literal dependency closures; no office conversion or
recalculation produces these goldens. Sixteen exclusions are recorded explicitly.
The [independent reproduction](native-reproduction.json) refetched and hashed
every raw input, regenerated identical JSON, and removed its temporary tree.

The frozen implementation passes all seven isolated gates: 1,348 ODS tests
(zero failures or ignored tests), all-target Clippy with warnings denied,
rustdoc with warnings denied, crate formatting, selected-file formatting,
crate boundaries, and diff whitespace checks. The [gate receipts](gates/)
retain the dependency lock, selected source freeze, complete compiled source
closure, commands, output logs, and unchanged before/after hashes.
The [semantic review](review.md) and [resource/cache review](resource-review.md)
accept the frozen source with no remaining correctness or resource blocker.

The final [performance report](performance/results/performance-report.md) contains
390 baseline and 3,000 candidate samples: 15 fresh processes per case and phase,
with three warmups. Matched allocation counts, requested bytes, and peak live
bytes are unchanged. Unrounded raw median time shifts range from −5.46% to
+1.90%; RSS shifts range from −4.09% to +3.92%. These are observed single-host
measurements, with no causal or cross-platform speedup claim. Nested AVERAGE,
COUNTA, and COUNTBLANK read 512, 2,048, and 8,192 cells at 64, 256, and 1,024
fixture rows, respectively: one outer IF scan plus one cached inner scan.

From the repository root, verify all retained receipts with
`python3 docs/report/spec-gap-validation-evidence/ods-formula-statistical-reducers/verify.py`.
The verifier regenerates the exact oracle, checks native reproduction and frozen
source identity, verifies all seven gate logs, checks every retained performance
file, and recomputes medians and read bounds from the raw samples. The selected
source files must match their recorded hashes. See the [root receipt](root-verification.json)
and [cleanup receipt](cleanup.json). The native helper can separately refetch and
regenerate the pinned source evidence.
