# ODS formula value inspection and conversion

This batch implements sixteen OpenFormula §6.13 functions: `ERROR.TYPE`,
`ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`,
`ISNUMBER`, `ISODD`, `ISTEXT`, `N`, `NA`, `NUMBERVALUE`, `TYPE`, and `VALUE`.
Independent semantic and resource reviews accept the corrected implementation.
All seven isolated gates and the 107-observation oracle pass. Two complete
performance captures independently verify, with explicit latency/RSS review
flags accepted in the [batch disposition](completion.md).

The [contract](contract.md) records the normative rules and explicit host
choices. Both evaluator profiles share private raw-value kernels. The value
evaluator retains Empty, formula errors, borrowed text and reference descriptors
until the function applies its own conversion rules. Ordinary matrix arguments
are read one coordinate at a time. TYPE scans complete references without
retaining cells; N keeps scalar intersection and the first inline-array element.
Typed provider, cancellation, resource and source-version failures propagate
through the existing evaluation boundary.

VALUE uses a fixed en_US grammar for numbers, currency, percentages, mixed
fractions, dates, times and datetimes. Calendar conversion shares the existing
TEXT helpers. The epoch is 1899-12-30, the calendar has no fictitious 1900 leap
day, and two-digit years use the 1930–2029 window. NUMBERVALUE applies explicit
decimal/group separators and the ordered OpenFormula normalization rules.
Long grouped numeric input uses budgeted scratch storage and exact decimal
conversion. The batch adds no locale, clock or recalculation provider.

The independent [Python oracle](inspection_oracle.py) generates the retained
[goldens](inspection-goldens.json), consumed by a Rust integration test. Native
observations and their reproducible fixture are retained under [native](native/).
Native differences remain explicit profile differences, not oracle agreement.

Preparation documents are retained byte-for-byte as hashed capture inputs.
Their provisional status statements are historical: `integration-plan.md` and
`performance/README.md` predate acceptance, and `native/README.md` describes the
initial 93-observation oracle. The current oracle has 107 observations and uses
contract SHA-256 `f227e861c55e3360d60d41a42ee7f101ba3c29bbc00d25d56cc01d80c3cd7923`.
The current [native provenance](native/provenance.json), goldens, verifier
results and [batch disposition](completion.md) supersede those provisional
counts and status statements.

[Gate scripts](gates/README.md) stage a declared source closure into a clean
baseline checkout and validate command logs and source hashes. They use the
retained gate Cargo.lock, whose identity differs from the ambient workspace
lock. The [performance harness](performance/) compares matched existing
functions and separately measures new inspection functions, reference reads,
matrix results and typed refusals. `verify.py` validates the retained evidence
and fails closed when required final receipts are missing.

The initial gate attempt was invalidated after a concurrent script self-check
removed the isolated lockfile. The frozen lock was restored and all seven gates
were rerun. The accepted rerun has identical before/after source manifests;
`gates/invalidated-attempt.json` records the discarded attempt.

Source review subsequently found and corrected signed-exponent/date ambiguity,
fractional-second digit order, omitted scalar separators, and premature
projection of computed Any arguments. The independent oracle and focused
regressions cover these cases. The earlier candidate's successful gates do not
serve as validation for the corrected snapshot.

This batch does not complete OpenFormula §6.13. Reference/host metadata
functions, configurable locale/date settings, dependency recalculation and
workbook cache publication remain outside its scope.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-value-inspection/verify.py`
from the repository root to recheck retained evidence after temporary checkout
cleanup. This checks the committed baseline source, candidate gate manifests,
raw measurements and both complete captures. Supplying `--candidate-root` and
`--candidate-freeze` additionally supports verification against a live staged
checkout. The retained [root receipt](root-verification.json) records the final
result.
