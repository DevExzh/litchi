# ODS formula tokenizer performance and allocation failures

This batch follows `b88341bf4`. The preceding reference-family implementation
added bounded ODF 1.4 Part 4 §5.8 metadata, but its paired microbenchmark recorded
18.2–29.0% higher latency for ordinary function/cell workloads than the earlier
parser. The audit's remaining formula grammar work makes this common tokenizer
path relevant to both existing and future supported formulas.

The scope is to reduce redundant tokenization work without changing public
reference types, grammar support, retained formula text, or finite limits.
Speculative cell recognition must propagate allocation/resource failures;
those failures are not evidence that an input should be retried as a name.

ADRs 0001, 0005, 0006, and 0008 require correctness and bounded resources before
speed, measured results, and a buildable verified increment. Formula ownership
remains in `litchi-ods` under ADRs 0002/0023/0024. No evaluator, external I/O,
workbook resolution, or package mutation is introduced.

The baseline is the committed reference parser, and the same harness measures
ordinary formulas and all newly supported reference forms on both sides.
The compact scanner recognizes complete function/cell candidates before copying
components. The legacy path also delays component copies and propagates typed
allocation errors. Space-bearing sheet names and absolute references retain
the compatibility path. Reference-parser changes are code-generation hints only;
the copy helpers still have out-of-line symbols in the captured release binary.

All five package gates passed on unchanged source: 815 tests across 48 targets,
Clippy with warnings denied, rustdoc with warnings denied, doctests, and formatting.
An isolated checkout passed another 69 focused tests. The allocation regression
fails against the baseline and passes for compact and legacy cell paths here.
See [boundary review](spec-review.md) and [gate receipts](gates/results.json).

The [paired performance report](performance/report.md) records 56 lanes.
Common function/cell inputs improved by 3.6–4.8% in measured p50 latency.
This is a mixed result: long colon-bearing sheet names regressed 14.7–20.0%,
and the over-limit IRI refusal regressed 32.8%. These observations do not establish
a general parser or package speedup. The compatibility scanner retains its
legacy quadratic worst case for long space-separated cell-like input.
Allocation totals and process RSS are reported separately from latency.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-tokenizer-performance/verify.py`
from the repository to verify the source-bound gates, raw results, binary digest
receipts, candidate patch replay, and artifact manifest. The small retained
harness and text evidence permit reproduction; temporary checkouts, build targets,
and saved executable copies are removed after validation. Root cleanup is recorded
in [root-temp-cleanup.json](gates/root-temp-cleanup.json).
