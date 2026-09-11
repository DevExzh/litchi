# Change 0507 evidence

The shipped candidate is `candidate.patch`, two private groups of eight ODG
attribute-value requests. The [change record](../../changes/0507-odg-attribute-value-batches.md)
contains individual results, allocation/instruction attribution and limits.

`capture.json`, `summary.json`, `before/` and `after/` are the primary serial
A1/B1/B2/A2 comparison: 16 children, 25 warmups and 200 samples per child.
The four `before/*-pilot.json` files are a separate 400-sample pilot. No paired
measured latency/throughput/RSS adverse change exceeds 5% in this capture.

For reproduction, build the unchanged locked 0502 probe at the clean base
revision with the command in `environment.json`, freeze the executable, apply
`candidate.patch`, then rebuild/freeze the candidate. Run with Python on one
shell line: `capture.py --before /absolute/before --after /absolute/after
--output /absolute/fresh-results`, then `analyze.py --root /absolute/fresh-results`.
CPU 2 and /usr/bin/time are required. The analyzer imports the retained 0502
bootstrap helper relative to its repository location. Do not overwrite these
reports. The source manifest identifies inputs and byte-identical final rebuild.

Callgrind and heaptrack files contain raw profiles and summaries. The default
profile corpus is metadata-large; `heap-plain-*` uses plain-large. Each child
has zero warmups, a preflight open and one timed sample, plus setup and hashing.
Counts therefore cover the whole child. `profiling.json` / `profiling-after.json`
and environment JSON record commands and tool versions. The fresh
`hardware-counter-probe` receipt documents the unavailable hardware counters.
Profiler RSS is not pooled with uninstrumented /usr/bin/time results.

`gates.py` reproduces the scoped checks with explicit target/output paths.
Its optional `--without-boundaries` flag permits the independent boundary check
to be recorded separately; root used that during this batch. The final
`gates.json` includes both outcomes. An initial test-only lifetime compile error
is retained in the `initial-*` logs; the full corrected suite passes 124 tests.

`adr-review.md` and the mechanically verified `adr-manifest.json` record
architecture inputs and obligations. The test-source manifest identifies final
ODG Rust sources. `cleanup.json` records removal of only batch-owned scratch;
`custody.json` hashes every other evidence file. This bundle does not claim a
full-workspace, fuzz, sanitizer, live Office or strict-registry run.
