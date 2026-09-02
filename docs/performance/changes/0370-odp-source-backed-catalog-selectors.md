# Change 0370: ODP source-backed catalog selectors

## Scope

Change 0370 adds three opt-in selectors to the performance harness:
`odp_source_backed_catalog_open`, `odp_source_backed_catalog_list`, and
`odp_source_backed_catalog_query`. It changes no production code, CRUD API,
format behavior, or default benchmark matrix. The selectable registry is
**407** and the default remains **36 cases / 198 rows**.

## Corpus and timing

The fixed media-rich ODP corpus contains 12 slides and 13 archive members,
including eight deterministic 2 MiB `Pictures/*` members. The archive is
16,785,912 bytes with SHA-256
`661ae80396d4eda673d35e45d208443cc359052e4b9b27fed0ba6681602a913a`.

The selectors have separate preparation and measurement scopes:

- `odp_source_backed_catalog_open` times fresh source-backed catalog
  construction.
- `odp_source_backed_catalog_list` times only `catalog()` after the owner is
  prepared.
- `odp_source_backed_catalog_query` times only the selected slide projection
  at query index 6 after the owner and index are prepared.

Semantic digests, topology, source-replay, and media-locality checks are
outside the timed regions. The `Pictures/*` members are retained as a
locality guard and are not read by the list projection after preparation.

## Descriptive control

The retained [machine-readable control report](../results/odp-source-catalog-0370-control.json)
uses CPU 2, 30 warmups, and 500 samples per selector. It was built from dirty
revision `f35486fb7085bb128eb89a4d2e9edd3ad1065f02`; the binary SHA-256 is
`08594839ede39d7f2ed0c143d818e41de0b7cdb77bc92fbcdd2a96083ca9966a`.

| Selector | p50 (ns) | mean (ns) | p95 (ns) | p99 (ns) |
| --- | ---: | ---: | ---: | ---: |
| open | 57,538 | 61,057.616 | 76,884 | 88,020 |
| list | 31 | 63.854 | 161 | 200 |
| query, index 6 | 60,062 | 64,323.154 | 83,354 | 101,659 |

These values are descriptive current-control observations only. The source is
dirty and there is no clean A/B comparison, so no speedup or hotspot ranking
is inferred.

## Validation and claim boundary

Strict Clippy exposed only 23 preexisting unrelated diagnostics; the scoped
allow-list Clippy run passed. The focused selector test passed `1/1` and the
selectable enumeration test passed `1/1`. No full suite was run.

`performance_claim: none`; `claim_authorized: false`. No latency, RSS,
allocation, physical-I/O, decompression, cold-cache, fixed-memory,
throughput, or OOM-prevention claim is made. This batch provides measurement
coverage for a future controlled ODP catalog experiment only.
