# Change 0269: XLSX repeated-store cache ABBA

Date: 2026-08-24

Status: LANDED on all eight accepted latency cells; this is a narrow
latency-only XLSX semantic-query claim and makes no broad XLSX claim

Claim registry ID: `claim-0269-xlsx-repeated-store-cache`

## Evidence-package identity

The checked-in ABBA package is
`docs/performance/results/0269-xlsx-repeated-store-cache-abba-20260824/`.
Its manifest is
`0269-xlsx-repeated-store-cache-abba-manifest.json` with SHA-256
`a1c9b662e9d84886d69130ef054c2cc4b1ac55952fd97fb48df1f8a6a96226e4`.
The `summary.json` file SHA-256 is
`6200fa280525d370d6d8375d750def700c011c7d646fddfbddec642afce8d317`; its
canonical SHA-256 is
`6a6697f01ac678efbc8c8466bc8f32a5ecef99700b1492437ab081f87bdea4ae`.
The package is schema version 1, has change ID
`0269-xlsx-repeated-store-cache-abba`, and contains two result rows from
`litchi-perf-abba-summary` 0.1.0.

The retained package members have these identities:

| Member | File SHA-256 | Canonical SHA-256 |
|---|---|---|
| `0269-xlsx-repeated-store-cache-abba-manifest.json` | `a1c9b662e9d84886d69130ef054c2cc4b1ac55952fd97fb48df1f8a6a96226e4` | — |
| `summary.json` | `6200fa280525d370d6d8375d750def700c011c7d646fddfbddec642afce8d317` | `6a6697f01ac678efbc8c8466bc8f32a5ecef99700b1492437ab081f87bdea4ae` |
| `a1.json.zst` | `25a158ea84372497f69584872016b8592a070ba39867b34828ee104c124fff77` | `2ecf8b594ad661e1fc69f02b97a0babef7e82fa7fa56fad455e2ccd327d5ce42` |
| `b1.json.zst` | `efde47d7f6ec9d4b8e7bf2d3747c36f5d98a564919fd56045156f3bf118792dd` | `5bead83a5593dddf5db627568710a7a5c50b2f6ac8a9877be4e74cd2e60561f3` |
| `b2.json.zst` | `b98ddfc8a96d145b50a5543d14196b5ffed2c8a8f228c89ec8e3f601e48b6b94` | `ed787f95bdd75315403435acbddbf3389cf7f08abeb723ab900b32d1406f0a5d` |
| `a2.json.zst` | `b119ee9319818da8fcd58100b483b0645a4a4e62b58200e35a530546c8e5f722` | `149b572bbb807c2d893077322946c10ee82de3915f5ddeb52a90cfedf2d556b7` |

The matched control revision is
`18633404d27bc4c442c09915972e7655cdae813b` (A1/A2); the candidate revision
is `8a0ca40b1a9d77a9494c74c0cdca38dd61ee68b1` (B1/B2). The control and
candidate release binaries are distinct, with SHA-256
`7dbb18227a97b228337b25518f5219315e239eca6e5b07bcb21582db521acfc4` and
`89594ae04dcce5d016216296a53409db79a1cd25acb2c82d1fa3ef5de13bcf6c`.

## Scope and pinned corpora

The two primary selectors use the existing source-backed worksheet store and
the exact semantic-query timing boundary:

| Selector | Corpus | Archive SHA-256 |
|---|---|---|
| `xlsx_source_repeated_store_medium` | `xlsx-source-repeated-store-medium` — four 48×48 worksheets, 9,216 scalar entries | `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` |
| `xlsx_source_repeated_store_oversized` | `xlsx-source-repeated-store-oversized` — four 48×48 worksheets with an oversized selected worksheet | `3cf797e44ef51189a4b62d040cf39ff2af670ebd909c6e806f387b51e72ecfec` |

Both corpora use generator
`litchi-xlsx-source-repeated-store-corpus-v1`, 17 ZIP members, selected
`xl/worksheets/sheet1.xml`, and target `Sheet1!A1` with one-byte scalar
payload. The medium archive is 4,226,429 bytes and its selected worksheet is
63,294 uncompressed bytes. The oversized archive is 4,236,114 bytes and its
selected worksheet is 8,389,041 uncompressed bytes. These exact generator,
shape, package-format, member, target, and archive identities are bound in the
registry and package.

## Protocol and exact latency result

The release `litchi-perf-baseline` harness ran A1 control, B1 candidate, B2
candidate, A2 control on CPU 2 with one execution worker, 20 warmups, and 500
retained samples per selector and leg. Every sample used a fresh child and a
warm filesystem-root selection. Each timed operation repeats the semantic
queries `cell`, `cells`, `visit`, and `stored_extent` eight times in that order.
The timing scope is exactly
`semantic_query_only; explicit PartData reacquisition excluded`. The ABBA
drift ceilings are 5%/5%/10%/15% for p50/mean/p95/p99; positive values in the
table mean lower candidate elapsed time.

| Selector | Statistic | A1→B1 | A2→B2 | Verdict |
|---|---|---:|---:|---|
| `xlsx_source_repeated_store_medium` | p50 | +56.239601% | +56.510587% | accepted |
| same | mean | +56.294105% | +56.514965% | accepted |
| same | p95 | +56.781541% | +55.426259% | accepted |
| same | p99 | +50.786266% | +52.138506% | accepted |
| `xlsx_source_repeated_store_oversized` | p50 | +99.892285% | +99.891347% | accepted |
| same | mean | +99.889026% | +99.888001% | accepted |
| same | p95 | +99.866426% | +99.860963% | accepted |
| same | p99 | +99.848674% | +99.846436% | accepted |

All eight cells are accepted and there are zero adverse-both cells. The
strict package recomputes all elapsed statistics from the four raw reports;
operation-metrics identities are equal within each row. Output, sink, and
source identities are consistently absent because this semantic-query-only
scope does not publish an output or claim physical/source-I/O identity.

## Production boundary and exclusions

The candidate's repeated-store cache change is landed in the current branch.
This registry entry is latency-only: it has no resource guardrail and claims
no allocation count, peak memory, RSS, physical I/O, device traffic,
decompression, cold-cache behavior, throughput, storage-media behavior,
publication/save latency, or real-producer breadth. The two structural
reacquisition-control selectors from change 0267 are not part of this claim;
their elapsed/query vectors remain structural cache/read evidence and are not
candidate latency comparators. This result does not generalize beyond the two
listed semantic-query selectors and pinned XLSX/OPC/ZIP corpora.
