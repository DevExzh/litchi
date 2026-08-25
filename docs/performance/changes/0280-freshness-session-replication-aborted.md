# Change 0280: freshness-session replication aborted

**Date:** 2026-08-25
**Status:** Aborted before smoke
**Performance claim:** none
**Retained samples:** 0

## Decision

Change 0280 was predeclared as an independent replication of the unchanged
change 0279 candidate. Retention required an independently passing 24,000-row
run and a pooled 0279+0280 check. Exact reproduction of both pinned binaries
was a mandatory pre-smoke identity gate.

The control rebuilt byte-identically. The candidate did not. Collection
therefore stopped before semantic smoke, timing, or statistics. No 0280
performance evidence exists, and the unchanged operation-scoped freshness
session remains rejected.

## Frozen protocol

The protocol retained the change 0279 corpus, selectors, order, CPU affinity,
warmups, counters, and unqualified 5% p50/mean/p95/p99 same-side drift gate.
It increased retained samples to 1,000 per cell per leg and required an
untrimmed pooled check over 1,500 samples per cell per leg.

The fixed identities were:

| Side | Revision | Expected bytes | Expected SHA-256 |
|---|---|---:|---|
| Control | 9adbf9ee527b4ccd2e3494c824c761973633ea14 | 8,416,176 | 394d2e56184bea77dc3b0682f25712739b819c1b8e02aa301d2ab2a291c9bf82 |
| Candidate | 0b0e1ed02c2fe4a262826a25a1f64b533bd0e686 | 8,412,832 | a4edd3969c3114d15b6f773b2e9493e80bf0d4d117e627f269e1d37a23c4b562 |

## Build identity result

The two release builds were initially launched concurrently in isolated
targets. The host killed the candidate compilation for memory pressure.
The candidate build was then resumed alone with CARGO_BUILD_JOBS=2. Future
Cargo build/test phases on this host must be sequential and job-limited.

| Side | Actual bytes | Actual SHA-256 | Exact match |
|---|---:|---|---|
| Control | 8,416,176 | 394d2e56184bea77dc3b0682f25712739b819c1b8e02aa301d2ab2a291c9bf82 | yes |
| Candidate | 8,413,200 | d95c6aa41fb4811fa30ef8dc3f11d77f0205278b48b7b2c8baa85d3b395cae6e | no |

The record does not infer why the completed candidate binary differed. It only
applies the predeclared identity rule: a mismatched binary is not a replication
binary.

## Claim boundary

There was no smoke run and no retained sample. Change 0280 cannot be used for
a latency, counter, semantic, tail, repeatability, or combined-evidence claim.
Change 0279 remains the only evidence for the candidate, and its strict verdict
remains rejected.

Do not rebuild and retry the unchanged candidate as another replication. The
next performance batch must move to a different bounded design or hotspot.

## Retained evidence

Evidence:
docs/performance/results/0280-freshness-session-replication-aborted-20260825/

The package contains the frozen protocol, the expected/actual identity result,
the zero-sample summary, and the artifact manifest.

Manifest SHA-256: e3558b66f455c11f1a92026b900fcdf28634445ea73e58593b01602ef0672c14
Manifest bytes: 581
