# OpenFormula radix evaluation

This batch adds the fourteen OpenDocument 1.4 Part 4 §6.19 radix converters to
`codec::formula::evaluation`: `BASE`, `DECIMAL`, and all twelve directed
binary/octal/decimal/hexadecimal conversions. It advances §10 of the
[specification audit](../../spec-gap-audit.md). Roman numerals, other function
families, references, arrays, workbook recalculation and cached-result publication
remain open.

## API and profile

Call the existing `evaluate_scalar` or `evaluate_scalar_with_context` with an
immutable parsed `Expression`. No public settings, dependencies, I/O or global
runtime are introduced. Formula errors are values; unsupported capabilities,
cancellation, work/storage limits and allocation failures remain typed failures.
All present arguments are eager in source order, including malformed arity calls.
Lazy surrounding handlers retain their existing branch-selection rules.

`BASE` uses uppercase ASCII digits for radices 2 through 36. Its Integer
parameters use Number conversion and truncate toward zero, an explicit choice
under §6.3.6. Padding is limited by the caller's text, work and memory budgets.
`DECIMAL` accumulates an exact unsigned integer and rounds once to nearest-even
`f64`. A fixed 1024-bit magnitude supports the full finite integer `f64` domain
without a heap big-integer dependency. Rounded nonfinite values return `#NUM!`.
Leading ASCII spaces/tabs, hexadecimal X/0X prefixes and H suffixes, and a binary
B suffix follow the specific grammar; other whitespace and signs are rejected.

The twelve direct converters use signed 10-, 30- and 40-bit binary/octal/hex
widths. Number X must already be integral; fractional X returns `#VALUE!`.
Numeric digit spellings are interpreted in the source radix. Decimal source
text permits a leading minus and arbitrarily many budgeted leading zeroes;
binary/octal/hex source text is restricted to at most ten digits. Malformed
spellings return `#VALUE!`; values outside the selected signed domain return
`#NUM!`. Logical values convert to zero or one.

Optional Digits truncates toward zero and accepts 0 through 10. Zero or a width
shorter than the natural result keeps the natural result. Negative values use
exactly ten sign-extended digits and ignore the evaluated Digits value, while
formula errors and typed failures still propagate. An explicit missing argument
is a Value error. The detailed normative choices are in [spec-review.md](spec-review.md).

## Validation

The frozen source passes **915 tests across 55 Cargo targets**, including twelve
focused radix integration groups, plus **two executed doctests**. All-feature
all-target tests, warning-denied Clippy and rustdoc, formatting, and crate-boundary
checks pass. Gate receipts bind 214 exact source inputs before and after.

The comparable corpus has 141 rows per revision; the candidate-only radix corpus
has 159 rows. All statuses and semantic preflights pass. Existing workloads retain
identical allocation counts, requested bytes, peak tracked heap and result checksums.
The 34 initial latency/RSS flags were repeated in four interleaved pairs (272 rows).
None retained a median-latency regression above 5%; seven retained tail-latency flags,
and one evaluation RSS lane retained a 5.48% increase. These remain disclosed review
limitations, not an overall performance improvement claim. Five serial hardware-counter
captures cover a common text workload, full-range BASE/DECIMAL, and large padding.
See [performance/report.md](performance/report.md) for individual measurements.

The initial benchmark draft was corrected before final candidate capture: stale
bitwise baseline refusals, incorrect fixed-width range/fraction expectations, an
over-limit concatenation, and an unwired cancellation case. The retained baseline
was rebuilt using the identical final harness.

## Architecture and cleanup

The private outlined handler adds no evaluator frame variants, dependencies,
public settings, cache state or I/O. Fixed stack arithmetic avoids heap big-integer
allocation; output uses fallible exact reservation retained by the returned scalar.
The caller's execution context charges work and storage and checks cancellation.
This follows ADR 0001's API layering, ADR 0002's crate ownership, ADR 0003's immutable
source/explicit operation model, ADR 0004's typed outcomes, and ADR 0005's bounded
memory and measured performance requirements.

After source, patch replay, profile, counter and binary provenance verification,
the isolated worktree, target and TMP directory were removed, reclaiming
**1,571,692,544 allocated bytes**. Unique scratch bytes were hash-verified in a
recovery archive outside tmpfs before removal. Both saved benchmark executables
were rehashed and removed after capture; receipts retain their identities.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-radix-evaluation/verify.py`
from this checkout to verify the retained artifacts, exact patch replay, source
hashes, gates, profiles, repeats, counters and cleanup receipts.
