# 0802 results review

This is a bounded review of the terminal packet in this directory. I checked
`summary.md`, `analysis.json`, `decision.json`, the selected raw native JSON
files, the quality receipts, and `execution-notes.md`; I did not rerun a build,
test, native capture, or profiler.

The terminal decision does not advance this candidate to workflow trials
(`advance_to_workflow_trials: false`) and does not adopt it in production.
The required `distinct-1` and `distinct-2` consume cases pass their benefit
rule, and `distinct-0` also improves, but four protected syntax consume cases
veto the candidate. The baseline therefore remains the production reference.

## Raw paired p50 spot check

For each selected raw JSON file, I verified 30 elapsed-time samples and
recomputed the nearest-rank p50 (sample index 14), then recomputed the median
of the six after/before block ratios. Values below are in nanoseconds. The
official bootstrap intervals are included for comparison; the recomputed
medians match the recorded analysis.

| case and mode | before block p50s | after block p50s | median ratio | official 95% CI |
| --- | --- | --- | ---: | --- |
| `distinct-0` consume (empty) | `[36430, 36330, 36340, 36330, 36370, 36430]` | `[24561, 24590, 24550, 24550, 24530, 24570]` | `0.675011` | `[0.674321, 0.676301]` |
| `distinct-1` consume | `[110730, 109051, 109801, 109890, 109761, 110290]` | `[87771, 87731, 87791, 87691, 87721, 88691]` | `0.799373` | `[0.795323, 0.804328]` |
| `distinct-2` consume | `[172911, 173591, 172901, 172991, 179961, 172921]` | `[160611, 160501, 162441, 160611, 160761, 162361]` | `0.928650` | `[0.908952, 0.939217]` |
| `syntax-flag-after-0` consume | `[72771, 78140, 72700, 76470, 74541, 76641]` | `[96430, 96221, 94930, 96230, 96470, 96291]` | `1.276295` | `[1.243891, 1.315446]` |
| `syntax-equals-value-after-0` consume | `[72760, 72860, 72800, 72740, 76630, 74520]` | `[97850, 99151, 98230, 98730, 97381, 98470]` | `1.347073` | `[1.296092, 1.359071]` |
| `syntax-flag-after-2` consume | `[212332, 212051, 213101, 212361, 212671, 219351]` | `[229651, 228301, 229491, 227761, 229452, 229431]` | `1.076772` | `[1.059236, 1.080236]` |
| `syntax-equals-value-after-2` consume | `[216212, 213831, 212321, 213641, 212631, 218771]` | `[228071, 228721, 229901, 229851, 230222, 229901]` | `1.072755` | `[1.052862, 1.082765]` |

The four syntax rows satisfy the frozen protected veto rule: ratio above
`1.05` and bootstrap lower endpoint above `1.0`. They are the four entries in
`decision.json`'s `protected_consume_regressions`, rather than a conclusion
inferred from a noisy individual sample.

## Coverage and diagnostic limits

The native packet covers 39 cases, construct and consume modes, six paired
blocks, and 30 samples per block: 936 native children and 28,080 elapsed
samples. The analysis records 58 diagnostic flags: all 39 construct groups
and 19 consume groups. The empty construct row is already a material
diagnostic (`distinct-0` ratio `3.177667`, or `+217.767%`); the other
construct rows have the same direction and are not hidden by the consume
benefits. Process p50 spread exceeds 5% in 66 of the 156 process groups
(39 cases × 2 modes × 2 legs), so the medians should remain diagnostic rather
than being presented as a general latency claim.

The candidate iterator measurement is 120 bytes for the baseline and 128
bytes for the candidate. The counter packet contains 312 positive Callgrind
dumps and 312 empty termination dumps, with 312 qualified owners and no
qualification failures. The counter values are guest diagnostics; allocator
matches are a lexical function-name census, not allocator API counts, and
inclusive descendant values overlap. Those records do not establish a public
workflow speedup or a construction-time resource claim.

## Quality and retained failure

The final helper quality matrix covers five minimal mirror crates: 70 baseline
tests and 100 candidate tests passed, with no failed or ignored tests. The
final baseline and candidate Clippy gates passed, and the direct probe receipt
records successful rustfmt, build, check, Clippy, and self-check commands.
This is helper-mirror coverage, not a full production-crate or workspace
validation.

One earlier candidate Clippy attempt is retained under `quality-failed-0/` and
`candidate-failed-0/`; it reported `clippy::question_mark` in
`OrderedCheck::next`. The final source changed only the equivalent
let-`Some`/return-`None` expression in the five helper copies, after which the
final gates passed. This was a quality-gate failure receipt, not a build or
semantic-test failure, and no production source was changed.

The frozen tested archive retains the stale trait-method documentation sentence
described in `execution-notes.md`. The implementation and module-level design
describe the bounded linear stage and conditional handoff correctly; the
sentence is a documented wording erratum and carries no performance or
adoption claim.
