# Rejected ASCII-name experiment

The pre-change AST harness at `9072e969e` exposed an avoidable binary search for
ASCII XML name characters. The isolated candidate returned ASCII alphabetic,
digit, and combining-class results before searching the unchanged historical
XML interval tables. `ascii-candidate.patch` records the exact attempted change.
An independent expanded-set check retained 34,514 Letter, 149 Digit, and 437
CombiningChar code points. The candidate patch also contained an exhaustive
ASCII unit test; it was not run before the experiment was rejected and is not
counted in the final validation totals.

All 24 baseline and candidate lanes produced identical outcomes, checksums,
allocation calls, requested bytes, and live-allocation peaks. ASCII-name p50
improved 51–69%. However, the subsequent balanced ABAB runs retained reference
p50 regressions of 7.1–9.9% across 64–4096 bytes and a 5.21% string-256 regression.
The string-4096 initial regression settled to 2.32% in ABAB.

The final production source reverts the entire experiment. The named workload
regressions trigger the review rule in `docs/GOAL.md`; the large name-only gain
does not conceal the reference regressions. No parser speedup is claimed for
this batch. The final scalar evaluator is measured separately.

A read-only binary review found unchanged sizes for all 22 reference symbols,
with 16 moving uniformly by -0x60 and IRI helpers moving by +0x20. Candidate
`.text` grew by 0x60, `scan_identifier` by 0x38, and `is_letter` by 0x20. Reference
parsing does not call the changed name predicates. These observations are
consistent with a code-placement/alignment effect; they do not prove a hardware
cause or establish that an inlining variant would avoid the regression.
