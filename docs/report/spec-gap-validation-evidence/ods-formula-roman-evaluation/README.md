# OpenFormula Roman numeral conversion

This batch adds `ROMAN` and `ARABIC` to the existing explicit scalar evaluator,
covering the remaining two OpenDocument 1.4 Part 4 §6.19 number-representation
conversion functions. The broader formula evaluation and recalculation gap in
[the audit](../../spec-gap-audit.md) remains open: other function families,
references, arrays, workbook dependency graphs and cached-result publication
are separate capabilities.

## Profile and algorithm

`ROMAN` accepts an Integer-converted number from 0 through 3999 and a format
from 0 through 4. Integer conversion uses the evaluator's deterministic
truncate-toward-zero profile. Logical format values map TRUE to format 0 and
FALSE to format 4 under the specification's legacy allowance. Zero returns
empty text, following the explicit zero rule and preserving the required
`ARABIC(ROMAN(0))` identity despite a later contradictory recommendation to
return the string `"0"`.

Formats 0 through 3 use the subtractive pairs permitted by their individual
tables. Format 1 allows V subtraction while format 2 forbids it, so output
lengths are not required to be monotonic across formats. Format 4 minimizes
symbol count under the standard's indirect-subtraction semantics with a fixed
six-ratio search over denomination residues. The independent numerical receipt
checks every value in the domain and verifies shortest lengths and ARABIC
round-trips.

`ARABIC` accepts ASCII `MDCLXVI` symbols case-insensitively, including the
empty string. A symbol subtracts when a larger symbol occurs anywhere to its
right; otherwise it adds. This accepts noncanonical spellings, including
spellings with negative results. Other characters and whitespace are formula
Value errors. Both functions retain eager source-order argument/error handling
and typed cancellation, resource-limit, allocation and unsupported-capability
failures.

## Scope and ownership

Call `evaluate_scalar` or `evaluate_scalar_with_context` with an immutable
parsed `Expression`. The implementation is owned by `litchi-ods`, adds no I/O,
network access, cache mutation, workbook reads, recalculation, public settings
or new dependencies, and leaves source formulas unchanged. References, arrays,
names, labels and other functions remain typed capability refusals in this
scalar profile.

Input scans and output construction use the caller's execution context, text,
work, storage and cancellation limits. The evaluator's string scanner checks
cancellation at 4096-byte windows and permits a doubled quote to straddle a
window boundary; the Roman regression covers concatenated literals across that
boundary. Format 4 uses fixed-size arrays and at most 64 candidate masks;
output construction uses checked lengths and a fallible reservation retained
by the returned owned text result.

## ADR compatibility

| Accepted ADR | This batch's evidence |
| --- | --- |
| 0001 priorities/API layers; 0004 semantic API | Existing scalar values and typed evaluator failures remain the public contract; no XML, package identifiers or new configuration leak into the API. |
| 0002 crate topology; 0023 ODF ownership | Implementation stays in the existing `litchi-ods` formula owner; no dependency edges or manifest changes. Crate-boundary gate passes. |
| 0003 snapshots/edits/patches | Evaluation borrows an immutable expression and performs no workbook or cache publication. Source spelling and force-marker preservation are tested. |
| 0005 I/O, memory and measured performance | Fixed search arrays, fallible retained output, charged work and bounded cancellation; exact-source before/after profiles and all regression flags remain visible. No hidden I/O, spill or parallelism. |
| 0006 validation/security/compatibility | ASCII validation, checked domains and lengths, eager formula-error order, distinct typed capability/resource failures, and explicit zero/format interpretation are covered. |
| 0008 verification | Scoped all-feature tests, Clippy, rustdoc, formatting, boundary checks, independent all-domain oracles and retained measurement provenance pass. No native application interoperability claim is inferred. |

## Validation

The final gate receipt records **926 tests across 56 Cargo targets**, with zero
failures or ignored tests, the eleven focused Roman integration groups, and **two
executed doctests**. All-feature all-target tests, warning-denied Clippy,
rustdoc, formatting and crate-boundary checks passed. The gate and source
digests are retained in [gates/results.json](gates/results.json); the detailed
normative and implementation reviews are [spec-review.md](spec-review.md) and
[implementation-review.md](implementation-review.md).

The independent numerical evidence is retained in
[numerical-review/receipt.json](numerical-review/receipt.json), with its
reproducible checker in [numerical-review/check_roman.py](numerical-review/check_roman.py).

## Performance

Final CPU 6 captures cover 159 comparable cases per revision and 123 candidate
Roman cases across parse, evaluate, and combined phases. Every result check
passes; comparable allocation counts, requested bytes, peak tracked heap and
retained output reservations are identical. This is a bounded formula profile,
not a workbook recalculation or package-wide memory claim.

All 34 initial latency/RSS flags received four interleaved A/B pairs (272 rows).
Two combined parse/evaluate p50 regressions remain: bitwise text coercion
**+7.60%** and the radix fractional-input error **+5.12%**. Eight lanes retain
p95/p99 flags, up to **+70.10%**; none retains an RSS increase above 5%.
These are explicit performance qualifications for the function-family
extension, not a no-regression claim. Follow-up should address the remaining
coercion/error paths under larger representative formula workloads without
trading away common-path performance. Fifteen timed batches do not establish
production tail percentiles on this shared host.

The string scanner removes a cancellation-threshold comparison per byte while
retaining a check per bounded window. Separate larger-batch hardware-counter
captures measure the 4096-byte UTF-8 lane at about 10.87 µs before and 10.09 µs
after; these measurements do not replace the combined-phase regression results.
Rejected outlining experiments and the original captures are retained under
`performance/exploration`. Final captures moved from CPU 2 to CPU 6 after an
unrelated job occupied CPU 2.

See [performance/report.md](performance/report.md) for commands, individual
results, uncertainty, counters and limitations. [verify.py](verify.py) checks
source/patch, gate, harness, binary, raw-result, repeat, counter, review and
artifact provenance, including final cleanup receipts.

## Cleanup

The finished worktree, Cargo target and TMPDIR were removed after final source,
review, capture and binary checks. Cleanup reclaimed **1,555,263,488 allocated
bytes** (about 1.45 GiB). Twenty-two unique scratch files were hash-verified in
a recovery archive outside `/tmp` and `/var/tmp`; the two final benchmark ELFs
were removed after hash verification. See [gates/cleanup.json](gates/cleanup.json)
and [gates/recovery-manifest.json](gates/recovery-manifest.json).
