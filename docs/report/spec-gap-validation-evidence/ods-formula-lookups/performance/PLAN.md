# Lookup performance capture plan

The profile compares baseline commit `635fd2e1348b621426b50909cbd5765c91837306`
with one frozen candidate. It is preparation only until the source gate,
contract review, and explicit root handoff are complete.

## Workload

The 33 matched controls are the 28 controls from the prior reference metadata
profile plus direct and projected `ROWS`, `ISREF`, and `ROW` descriptor
controls. They include both `reference-conditional-256x4-sumifs` phases so the
existing SUMIFS signal remains visible. The candidate adds 88 cases across the
nine lookup functions. The exact source, shape, expected output, lane, and
read policy live in `case-matrix.json`; the Rust harness repeats those values
coordinate by coordinate during its untimed preflight.

The fixture provides sorted vertical tables in columns A:C and sorted
horizontal tables in rows 100:101 on Main, Data, Archive, and Hidden. Search
cases use 8, 64, and 256 element sizes. Exact MATCH/HLOOKUP/VLOOKUP cases reach
the end or miss, exposing linear traversal. Approximate cases use middle or
last keys and record the reviewed binary probe count. LOOKUP covers tall and
wide result orientation, literal text, a shape refusal, and a projected
consumer. INDEX and OFFSET have zero-read descriptors and selected or matrix
consumers. INDIRECT has A1, R1C1, sheet, range, invalid, lazy, and projected
dynamic-descriptor branches. The projected dynamic lanes use borrowed `C1`
selector text in `H1`; their 8/64/256 coordinate cases expect 10 times the
number of selected coordinates and two reads per coordinate (selector plus
result), exposing linear descriptor-probe scaling.

The cross-cutting cases include lazy CHOOSE, selected references, projected
CHOOSE, an invariant projected MATCH, position-sensitive array and nested MUNIT
MATCH keys, resource and cancellation failures, list/shape refusal, output
matrices, and the existing metadata descriptor controls. Read accounting comes from the
complete `expected_reference_reads.by_case` map; cancellation is the only
bounded preflight interval.

## Capture protocol

Use three warmups and fifteen fresh child samples in both `evaluate` and
`parse-evaluate`. The arithmetic is:

```text
baseline: 33 controls × 2 phases × 15 = 990 rows
candidate: (33 controls + 88 lookup cases) × 2 phases × 15 = 3,630 rows
total: 4,620 rows
```

The runner builds in isolated temporary targets using the retained gate lock,
performs a complete candidate preflight before baseline timing, and snapshots
the selected source closure and all profile inputs before and after capture.
Each timed child validates its one sample, balances allocator bytes, and emits
raw stdout/time receipts. `verify.py` checks exact result-independent accounting,
per-case reads, source identity, and the flat retained manifest. The independent
`root_performance_audit.py` uses the same matrix and exact normalized elapsed
field when it is run after raw receipts are retained.

Do not start timing from this preparation tree. Once root freezes the source,
verify the pinned contract SHA256, update any contract-reviewed read exceptions in the
matrix, run syntax/format checks, and invoke the command in `README.md` from
the repository root. Preserve any failed setup or preflight under a diagnostic
sibling and record its reason; do not rerun to select a more favorable metric.
