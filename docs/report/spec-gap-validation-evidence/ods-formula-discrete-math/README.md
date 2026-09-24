# ODS discrete mathematical function validation

This batch covers `COMBIN`, `COMBINA`, `FACT`, `FACTDOUBLE`, `GCD`, `LCM`,
`MULTINOMIAL`, `EVEN`, `ODD`, `DELTA`, and `GESTEP` in the read-only formula
evaluators. Formula caches remain inert. The resolver-free evaluator handles
scalar arguments; the value evaluator supplies array and reference behavior.

The [contract](contract.md) records the normative signatures, conversion and
domain rules, optional defaults, formula-error ordering, and finite-binary64
numeric profile. The distinction between NumberSequence and NumberSequenceList
matters for reference-list admission. References select numeric/error cells;
inline arrays follow the established element conversion profile.

The [numerical oracle](numeric_oracle.py) uses Python standard-library bigint
arithmetic over the exact values represented by binary64 inputs. Fractional multinomial
references use exact rational sums and independent factorial ratios. It retains
[361 observations](numeric-goldens.json), including 224 seeded cases. The
reference does not compute floating factorial ratios or reproduce production
numeric kernels. The Rust integration test consumes the same observations
through both evaluator APIs. Cases cover factorial overflow boundaries,
combinations with finite results despite overflowing factorials, huge integer
operands, repeated LCM operands, an overflowing LCM followed by zero, and
multinomial totals beyond binary64's consecutive-integer range.

All accepted numeric observations require exact final binary64 bit matches.
The data retains mathematical reference bits separately from explicit profile
errors for the strict `COMBINA` domain and unrepresentable odd results.
Fractional cases include a represented sum just below an integer boundary:
`MULTINOMIAL(1.4;0.6)` returns `1`, without first rounding the sum to `2`.

These numerical cases are targeted evidence, not a proof over every input.
The independent semantic tests cover conversions, references, shapes, resource
failures, and formula-error propagation separately.

The [native evidence](native/README.md) retains 77 cached observations across
all eleven functions from eleven exact LibreOffice sources at a pinned commit.
Fifteen explicit exclusions distinguish source dependencies and producer
choices from this profile. These are cached-value comparisons, not native
application execution or save/reopen tests. The independent
[reproduction receipt](native-reproduction.json) records successful source
downloads, hash checks, byte-for-byte receipt reproduction, and temporary-tree
cleanup.

The frozen isolated checkout passes [all seven gates](gates/results.json):
1,271 ODS tests with no failures or ignored tests, strict all-target Clippy,
warnings-denied rustdoc, complete crate formatting, selected-file formatting,
dependency boundaries, and diff checks. The [freeze](gates/freeze.json)
identifies nine Rust files, three included JSON inputs, and the retained lock.
The gate runner records the complete workspace Rust/Cargo/build source closure
and included ODS test inputs before and after the checks.

The new reducers share the existing reference traversal and projection cache.
Their mutually exclusive numeric states use enum storage. The exact integer
kernel has a `u64` GCD path and a fixed-width path for larger values; callers
admit bounded kernel work before execution. Generated domain errors retain
source order, including when a later text conversion also fails. Typed host
resource failures remain outside `IFERROR`.

The [independent review](review.md) accepts the frozen implementation and
documents the fixed-width bounds, GCD termination proof, exact fractional
conversion, source-order errors, and structural ReferenceList refusal without
provider reads.

The [performance report](performance/results/performance-report.md) retains
1,860 fresh-process samples: 330 baseline and 1,530 candidate measurements,
with three warmups and fifteen samples per phase group. All 22 matched control
groups have identical median allocation calls, requested bytes, peak live
bytes, and retained result budget. Raw elapsed-time median changes range from
−3.10% to +2.54%. Process RSS medians increase by 0.25% to 7.00%.
DSUM crosses the 5% RSS review trigger: +204 KiB for evaluation and +168 KiB
for parse/evaluation. The heap metrics remain unchanged; the process-level
cause is unresolved and this measured increase is retained explicitly.

The three projected reducers each retain 768, 3,072, and 12,288 provider reads
at 256, 1,024, and 4,096 elements. These measurements establish scaling for
the named fixtures. The previously documented large-literal-expression cache
classification limitation is unchanged; these cases do not prove a bound for
every expression tree.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-discrete-math/verify.py`
from the repository root to independently verify the retained evidence. The
[verifier](verify.py) reproduces the numerical dataset, checks gate/native
receipts and source identity, validates every raw performance row, requires
the complete 51-case matrix, and recomputes report medians and provider-read
bounds. Its [receipt](root-verification.json) passes all checks.
The frozen profiler verifier contains a case-name spelling error; a
[retained verification adapter](performance/results/verify_retained.py)
applies exactly that correction after checking the original script hash.
The [verification notes](performance/results/verification-notes.md) explain
the correction and RSS review. Capture inputs remain unchanged.
