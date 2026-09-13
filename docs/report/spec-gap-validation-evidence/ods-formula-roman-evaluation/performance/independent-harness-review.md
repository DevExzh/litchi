# Independent Roman harness review

Reviewed the retained harness statically on 2026-09-13. No build, binary, or
benchmark process was run for this review.

The corpus declares 53 comparable cases and 41 candidate-only Roman cases.
Each case is sent through `parse`, `evaluate`, and `parse-evaluate`, so one
full candidate run contains 282 rows. The capture wrapper runs the 53
comparable cases for both revisions and the 41 Roman cases for the candidate:
159 baseline rows, 282 candidate rows, and 441 rows total. Every row uses the
declared three warmups and 15 measured iterations. Scale cases use repeats of
64, 32, 8, or 2 for inputs of 64, 256, 1024, or 4096 units; fixed cases use
128 repeats.

The 41 Roman cases cover 15 format vectors (3888, 499, and 998 across formats
0 through 4), three uppercase/lowercase/indirect ARABIC inputs, zero and
truncation, both Logical format mappings, three formula-error cases, four
domain/empty-input cases, four typed refusal cases, four 64-to-4096-symbol
ARABIC scans, and four 64-to-4096-call concatenation expressions. Their
inputs and expected values/refusals are wired through `input`,
`expected_value`, and `expected_failure`; the candidate's format 2/3 `ID`
vectors match the checked-in ODF reading. The 4096-symbol scanner and
concatenation cases exercise input growth and owned output growth rather than
only fixed-size examples.

The three phases measure different operations. `parse` parses and drops a new
expression on every repeat. `evaluate` parses once before timing and evaluates
that immutable tree repeatedly. `parse-evaluate` parses and evaluates on every
repeat. A preflight executes once before warmups to validate the expected
scalar or typed refusal; timed iterations then record success/refusal counts,
checksums, and failure labels without repeating semantic assertions.

Context and parsed-tree setup are outside the timer. Allocator counters reset
after setup, `live_before` records the persistent baseline, and the evaluator
result is dropped on every iteration. The context remains alive through the
`live_after` read, so result leaks remain observable. `requested_bytes` and
`released_bytes` are process allocator totals in the timed region; a realloc
counts its new size as requested and its old size as released. `peak_live_delta`
uses the global live-byte counter, while `output_reserved_bytes` records the
largest retained evaluator result reservation in a sample.

Two scope limits are material when interpreting results. The binary's
`Instant` timing excludes process launch, setup, and preflight, whereas
`/usr/bin/time -v` RSS is process-wide and includes startup. Also, `run.py`
does not itself exit nonzero when a child row has a nonzero status; the capture
wrapper checks every raw row status and aborts on such a result. Direct runner
use should therefore inspect `raw.csv` and status sidecars.

Source hashes at review time:

* `run.py`: `ad256b7938c91da3fad58d25206d3e85dddd10c5a3d0bbb809ae1952406e5c3b`
* `src/main.rs`: `72d2ccc147fc04db683b64d426adafe2bde7243c9d47707b0e68651e3a4806c5`
