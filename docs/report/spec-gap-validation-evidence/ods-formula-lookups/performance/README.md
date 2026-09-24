# Lookup performance profile

This directory owns the preparation inputs for the nine-function lookup profile:
`ADDRESS`, `CHOOSE`, `HLOOKUP`, `INDEX`, `INDIRECT`, `LOOKUP`, `MATCH`,
`OFFSET`, and `VLOOKUP`. The matched baseline is Git commit
`635fd2e1348b621426b50909cbd5765c91837306`; its isolated `Cargo.lock` is the
retained gate lock identified by `run_profile.py`.

The case matrix has 33 matched controls and 88 lookup cases. The controls retain
the existing SUMIFS evaluate and parse-evaluate lanes, scalar/array controls,
and five direct/projected ROWS, ISREF, and ROW metadata controls. Candidate
capture therefore has 121 cases. Two phases, three warmups, and fifteen fresh
child samples produce 990 baseline rows and 3,630 candidate rows, 4,620 rows
in total.

The lookup lanes cover ADDRESS A1/R1C1 formatting, lazy and projected CHOOSE,
exact linear and sorted approximate search scaling, selected-value reads,
INDEX/OFFSET/INDIRECT descriptor construction, A1/R1C1 INDIRECT, matrix
outputs, shape/list refusals, resource limits, cancellation, and typed formula
errors. The projected MATCH lanes check reuse of an invariant direct search
descriptor, a position-sensitive array search key, and a nested MUNIT scalar
key whose two projected positions produce distinct results. The projected
dynamic INDIRECT lanes scale the same descriptor lookup across 8, 64, and 256
coordinates using H1=`"C1"` and C1=10; each coordinate is expected to read its
selector and selected result. Every case has an exact expected scalar or
matrix result in `case-matrix.json`; the matrix also contains a complete
`expected_reference_reads.by_case` map.

Descriptor construction and shape refusal lanes must report zero resolver cell
reads. Consumer lanes charge only the selected cell(s). Exact scans expose
their ordered linear read count. Sorted approximate lanes use the reviewed
per-case binary probe count, including five reads for the descending LOOKUP
lane (four key probes plus its selected result), while cancellation retains a
bounded 0..1 preflight interval and the fixture's exact audit count of one
first attempted read. The nested MUNIT MATCH lane reads G1/G2 at its two
projected positions, then performs one and two exact search probes, for five
reads and results 1/2. No global zero-read assumption is used.

`run_profile.py` records `elapsed_ns_p50 / repeat` as
`elapsed_ns_per_repeat_exact`; summary and audit code take medians and deltas
from that float before any integer display rounding. The retained manifest is a
flat relative-path to SHA256 object at `results/retained-files.json` and
excludes itself.

The runner remains capture-gated. The reviewed contract digest is pinned in
`run_profile.py`. Before timing, the root owner must stage the candidate freeze
and send the explicit frozen-source handoff. The preparation command after that
handoff is:

```text
python3 docs/report/spec-gap-validation-evidence/ods-formula-lookups/performance/run_profile.py \
  --candidate-root <frozen-candidate-checkout> \
  --candidate-freeze docs/report/spec-gap-validation-evidence/ods-formula-lookups/gates/freeze.json \
  --warmups 3 --samples 15
```

The command must run from the repository root. A prior output directory is
moved to a timestamped `diagnostic-*` sibling before a retry; completed raw
receipts are never silently overwritten. Host load, CPU affinity, and any
unrelated workload are recorded as capture context, and the profile does not
claim isolated-host timing.
