# Change 0272: OPC source-overlay multi-part matrix

Date: 2026-08-24

Status: benchmark-only exploratory evidence; `performance_claim: none`

## Scope and current counts

Three new opt-in selectors cover the changed, equal-payload no-op, and mixed
source-overlay multi-part paths. Each selector contributes nine records across
the Cartesian set of source sizes `2`, `8`, and `32` and payload shapes
`small`, `large`, and `media-incompressible`, for 27 opt-in records total.
The selectable matrix is now **392 names**. The default remains **36 cases /
198 records**.

## Evidence contract

The matrix records exact phase vectors and binds the source identity, cache
state/counters, sink status, and raw-record/order identities for each record.
The correctness oracles retain the exact comment metadata and ordering checks
alongside the raw ZIP-record checks. Fixed corpus, revision, binary, and
configuration identities are required, and the ABBA validator checks paired
ordering and identity before any comparison output is accepted.

This is evidence-boundary and correctness coverage only. The validator and
recorded oracles do not turn the matrix into a latency, allocation, RSS,
physical-I/O, decompression, copy, throughput, or production-performance
claim.

## Exploratory profile disposition

A dirty five-sample smoke/profile run recorded 2,110 `cycles` samples with no
lost samples. Within that profile, zlib `deflate_medium` accounted for
47.37% and `longest_match` for 17.46%. These numbers are prioritization
observations only: the worktree was dirty, the sample count was five, and the
profile is not retained evidence or a claim. Latency and resource conclusions
are explicitly withheld.

## Production boundary and open gap

No production optimization is adopted by this change because recompression is
required for the changed/mixed publication paths. Any future parallel or
compression-policy change requires an explicit execution context and scaling
evidence before it can be considered. The change adds no production API,
dependency, default case, claim-registry entry, or historical classification
update.
