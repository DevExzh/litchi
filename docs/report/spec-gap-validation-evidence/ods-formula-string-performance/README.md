# ODS formula literal memory bounds

The baseline is `cbc60f112`. Its literal decoder reserves the entire remaining
formula for every string token. Many short literals therefore retain quadratic
String capacity even though formula bytes and token count are finite. This batch
addresses that resource amplification before expanding the remaining expression
grammar in the specification-gap audit.

The format owner remains `litchi-ods` under ADRs 0002/0023/0024. Correctness,
lossless text retention and fallible allocation take priority over speed under
ADRs 0001/0005/0006. Decoded literal storage should be proportional to the decoded
content, with no extra heap scratch and with unchanged public token values and
formula limits. The bundled ODF 1.4 Part 4 §5.4 also excludes U+0000 from literal content.
This batch closes the existing NUL acceptance gap with a typed syntax refusal;
other literal characters retain their grammar. No evaluation or external I/O is
added. [Specification provenance](specification.json).

The syntax pass finds the closing quote and counts doubled-quote pairs. The
decoded byte length is the encoded content span minus that pair count. After
syntax succeeds, one fallible reservation admits exactly that decoded length;
empty literals retain zero capacity. Plain strings copy one UTF-8 span directly,
and escaped strings copy complete UTF-8 segments while collapsing quote pairs.
The sum of literal reservation requests is bounded by their combined encoded
content, rather than by the sum of all remaining formula suffixes. Original
formula text remains a separate retained allocation.

Unterminated strings and NUL are now refused before reserving literal storage.
Valid literal allocation failures retain `Error::Allocation` and the resource
`formula string literal`. Formula-byte admission and pre-token token-count
admission keep their existing order. Six regression groups cover retained
capacity, UTF-8/quote values, exact text, NUL and other controls, malformed quotes,
limits, and targeted allocation failure. On the baseline, the capacity and NUL
tests fail while the other four pass. The candidate passes all six.

The frozen source passed 828 tests across 50 targets plus Clippy with warnings
denied, warning-denied rustdoc, doctests, and formatting. See
[gates/results.json](gates/results.json) and [independent review](spec-review.md).
The [performance report](performance/report.md) records paired ordinary,
reference, and literal workloads with allocation totals and process RSS.
The final paired run reduces peak live allocation for 4,096 one-byte literals
from 34,488,323 to 937,987 bytes (97.3%), with parse latency improving from
341 to 310 microseconds. Empty and mixed UTF-8/escaped literal cases also use
less memory and run faster. The change has a measured cost: targeted A/B/A/B
runs show approximately 8–19% slower parsing for 64–1,024 one-byte literals,
13% for a single plain 64 KiB literal, and 51% for a dense doubled-quote
literal. These costs are accepted for bounded retained storage and normative
NUL rejection, not presented as a general tokenizer speedup. The dense quote
case performs a sizing pass before decoding; it encodes 32 KiB of output in
64 KiB of quote pairs. Ordinary reference cases remain close to baseline;
all individual regression flags and measurement limits are in the report.
These are in-memory parse measurements, not document CRUD measurements.

The harness checksum covers token count and original-text length; decoded-value
correctness is established by the regression tests and source review.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-string-performance/verify.py`
to check source provenance, gates, the baseline failure, raw metrics and artifact
digests. The small harness and evidence are retained; temporary checkouts, build
targets and executable copies are removed after capture. Complete expression
grammar, function arity, and evaluation remain open in the broader audit.
