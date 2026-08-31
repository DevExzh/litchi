# Change 0347: XLSX cell-values harness calculation-closure repair

Date: 2026-08-31

Status: harness repair and bounded evidence only; no production candidate

Performance claim: none

## Decision

The old clean control binary failed the source-backed XLSX cell CRUD oracle
with `XLSX source CRUD changed raw unselected ZIP member xl/workbook.xml`.
That member was intentionally rewritten by calculation invalidation: an
effective cell edit rewrites `calcPr`, and an existing calculation chain is
removed as a package closure. The harness now permits only that closure and
continues to require raw identity for every other ZIP member.

The repaired harness passes the bounded direct medium one-edit smoke for both
eager and source-backed legs with three warmups and 30 retained samples. A
24-row ABBA v1 smoke with zero warmups and one sample per case records
`failure_rows=0` and `complete_rows=true`. The timing gates are not passed and
the result is not claim-authorized: `timing_gates_passed=false`,
`claim_authorized=false`, and `performance_claim: none`.

No production optimization was made. Source planning and commit dominate the
source-backed direct path in the retained phase evidence; a shared
publication-copy design remains unproven and deferred.

## Direct smoke

The retained raw report is
[`direct-medium-one-edit.json`](../results/change-0347/direct-medium-one-edit.json).
It is copied unchanged from the bounded evidence directory. The report uses
Rust 1.95.0, CPU affinity 2, one execution worker, three warmups, 30 samples,
fresh process isolation per sample, and the fixed medium four-sheet numeric
corpus.

The corpus has 17 ZIP members, 9,216 scalar entries, 4,226,429 archive bytes,
and SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.
The four worksheet members are `sheet1.xml` through `sheet4.xml`; shared
strings are absent and styles are held in `xl/styles.xml`.

| Leg | p50 ns | Mean ns | p95 ns | p99 ns | Output SHA-256 |
| --- | ---: | ---: | ---: | ---: | --- |
| Eager | 9,028,668 | 9,206,709.166666668 | 10,386,013 | 10,455,471 | `9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1` |
| Source-backed | 8,021,234 | 8,103,566.1 | 8,572,825 | 8,916,088 | `9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1` |

The source-backed leg records 240 logical source reads and 4,232,733 logical
bytes per sample, 143 ordinary payload reads and 4,223,551 ordinary payload
bytes, three payload materializations, and maximum in-flight reads of one.
Its semantic SHA-256 is
`3cd21160d4f74fa0f097ab40be08e211b3e460cea788aa2b6705a55fdece07de` and its
untouched-member evidence is count 15 with SHA-256
`7105fcbce160328f666e69fcfd18da9e19fd71dd7b63961e7cddd29d5da1a17d`.

## ABBA smoke and identities

The ABBA smoke was run as a standalone 24-row v1 batch with
`warmup=0` and `samples=1`, one serialized child at a time. Its bounded status
is retained in [`summary.json`](../results/change-0347/summary.json); no
unbounded child output or target directory is copied.

The v1 numeric corpus pins these identities:

| Shape | Archive bytes | Archive SHA-256 | Untouched count: one-edit / one-percent / batch |
| --- | ---: | --- | ---: |
| `medium` | 4,226,429 | `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` | 15 / 12 / 12 |
| `dense-sparse` | 4,251,863 | `893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a` | 15 / 12 / 12 |

Only numeric cell-values workloads are claimed. Formula and date workloads are
explicitly excluded. The ABBA schema remains
`litchi.xlsx.cell-values-abba.v1`.

## Closure and safety gates

- Direct `calcPr` invalidation is checked by namespace-aware XML matching and exact flag validation.
- An existing calc-chain may remove only its exact workbook relationship, target part, and content-type Override; orphan, duplicate, foreign-prefix, suffix-near-miss, and chain-addition topologies fail closed.
- Workbook XML outside direct `calcPr`, all other part metadata, media payloads, ZIP local records, and central records remain byte-checked.
- Effective current workloads require the exact changed-member set; no-op requires an empty changed-member set and exact source bytes.
- The source semantic digest is compared with the eager typed semantic digest without weakening raw output or ABBA B1/B2 gates.
- XML normalization has finite byte, depth, token, namespace, and allocation bounds; duplicate normalized ZIP names are rejected.
- The host-memory-safe protocol uses one worker and serialized child processes. No parallel rebuild is part of this evidence.

## Focused source tests

The repaired source includes these focused tests; no separate test command was
run during this record pass:

- `current_no_chain_workbook_shape_has_exact_calc_pr_closure`
- `unexpected_untouched_workbook_mutation_is_not_normalized_away`
- `calc_chain_closure_removes_only_relationship_part_and_override`
- `calc_chain_relationship_matching_is_exact_for_both_canonical_uris`
- `duplicate_normalized_zip_member_names_are_rejected`
- `pinned_current_shape_identities_are_exact`
- `pinned_current_shape_untouched_counts_are_exact`

## Evidence limits

The old control failure and the repaired direct/ABBA status are evidence for
harness correctness and protocol attribution only. No before/after production
comparison, allocation/RSS measurement, physical-I/O claim, or general XLSX
CRUD claim follows. The shared publication-copy idea is deferred until a
separate serialized, resource-bounded experiment can isolate it.
