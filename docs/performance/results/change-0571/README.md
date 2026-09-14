# change-0571 evidence packet: one archive index per detect-and-open

Change record:
[`docs/performance/0571-ooxml-prepared-source.md`](../../0571-ooxml-prepared-source.md).
Disposition: retained. `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
|  `public-api.md` | The public API added by this change, verbatim, with the ADR constraints it has to satisfy. The measurements and the compliance argument are in the change record itself. |
| `gates.sh`, `gate*.txt` | The gate commands and their raw output. |

## Result

Archive constructions for a detect-then-open sequence fall from **2 to 1** for a
document, a presentation and a workbook. Reads on the package fall 47 to 24, 132
to 59 and 44 to 25. Paired medians over 3,000 iterations, twice, give savings of
35.9, 201.7 and 42.0 microseconds, which are 48.6%, 46.4% and 39.3% of the
two-call pattern — shares that reproduce change 0569's independent measurement
closely.

## Why this change exists rather than the documentation fix

Change 0567 proposed telling callers not to pre-detect. Change 0569 measured that
advice and rejected it: one of the named openers performs no detection at all,
the classification is computed and discarded so no facade error names the
detected format, and the pattern is unavoidable for any caller dispatching across
the three facade types. This packet is the fix that measurement pointed to.

## What is not here

No cold-cache, allocation, peak-RSS, throughput or multi-fixture-corpus capture.
Warm cache, one host, one fixture per format. The construction and read counts
are exact; the microsecond figures are microbenchmark medians, not profile
attributions. The scratch probes and their build output were removed after
capture.
