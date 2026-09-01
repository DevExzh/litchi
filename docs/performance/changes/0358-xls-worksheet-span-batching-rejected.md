# Change 0358: rejected XLS worksheet span batching

**Date:** 2026-09-01
**Status:** Rejected and reverted
**Performance claim:** none
**Retained samples:** 12,000

## Decision

The bounded worksheet-span candidate was rejected and reverted because its
predeclared serial ABBA protocol failed exactly five p99 gates. All 24 groups
completed, and every child passed schema, oracle, source, identity, and
semantic validation. The candidate's claim-bearing p50/mean directions passed,
but the unqualified p99 stability gates did not. No gate was narrowed and the
run was not repeated. Production and candidate tests were reverted; the new
OOM-bounded serial ABBA driver and its evidence are retained as reusable
infrastructure.

## Candidate under test

The candidate batched consecutive stateless XLS worksheet payload reads into
spans bounded by 64 KiB and 1,024 consecutive payloads. It introduced no CFB
API. The batch was intended to reduce repeated source reads and freshness
probes without changing worksheet semantics or the public XLS surface.

## Correctness evidence before measurement

While the candidate was present, the following checks passed:

- Python driver: `15/15`;
- worksheet-span checks: `9/9`;
- `source_backed`: `46/46`;
- `litchi-xls` library: `1021/1021`;
- CFB cursor checks: `7/7`;
- fragmentation checks: `9/9`.

## Strict ABBA protocol

The corpus was `test-data/ole/xls/ConditionalFormattingSamples.xls`,
1,402,368 bytes, SHA-256
`d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7`.
The control revision was
`7577e61224c0bd3d77b86ebb20e6d392d4f572af`; the candidate revision was
`7577e6122+source-84c669eeb6b2218cd7262398c5f966aa3781b0f826fc7fd0436791db5579e89e`.
The pinned control binary was 8,667,072 bytes with SHA-256
`397a5ba49047c3e362d1fe0f69810083ccda5154dbed0a169ebf6e09731acc1e`; the
candidate binary was 8,685,440 bytes with SHA-256
`9514f207734c553f0262fed4e6794c46748ea33a2dd6af4ebe401e602a6ef7c1`.

The six selectors were crossed with four A1/B1/B2/A2 legs and 500 retained
samples per cell, producing 12,000 retained samples. Collection used CPU 2,
20 warmups per fresh child, one child at a time, a 2 GiB child cap, no retries,
one sequential Cargo build lane, and an on-disk target. All 24 groups were
complete with no child, schema, oracle, source, or identity failure.

## Mechanism and descriptive result

For one-cell source-backed reads, the candidate changed the mechanism counters
by `+316` read bytes, `-79` reads, and `-158` version calls. The protocol's
claim-bearing file-source/one-cell p50/mean deltas were:

| Selector | A1 -> B1 p50 / mean | A2 -> B2 p50 / mean |
|---|---:|---:|
| FileSource one-cell | `+4.984845886382849%` / `+4.771328093073383%` | `+5.785582423178705%` / `+5.78027327071032%` |

These observations are descriptive only. The rejected result does not
establish a production latency, source-read, freshness, allocation, RSS,
physical-I/O, or broad XLS claim.

## Failed keep gate

Exactly five predeclared p99 gates exceeded the unqualified 5% limit:

| Selector | Gate and leg | Drift |
|---|---|---:|
| FileSource list | same-side A1 -> A2 | `+6.59542478684531%` |
| FileSource list | candidate-tail A2 -> B2 | `-6.916640348285569%` |
| FileSource one-cell | same-side B1 -> B2 | `+8.748517200474495%` |
| AtomicFile one-cell | same-side A1 -> A2 | `-5.699947129465672%` |
| AtomicFile one-cell | same-side B1 -> B2 | `+6.439283716879541%` |

The candidate is therefore rejected exactly as collected. No retrospective gate
narrowing or rerun is permitted; the production change and candidate tests
were reverted.

## Retained evidence and resource boundary

Evidence:
`docs/performance/results/0358-xls-source-span-abba-20260901/`

The approximately 7.4M evidence package retains the frozen protocol,
identities, oracle, all 24 groups, 12,000 normalized samples, comparisons,
summary, the five predeclared gate failures in `failures.log`, and artifact
manifest. Manifest SHA-256:
`c61f0be7ff8a0b894c04dd270203f4a64bd7a2f9d548c7dbf7bd071e1b31f4d2`.
Manifest bytes: `1175`.

The serial run used one on-disk target, one worktree, one Cargo build lane, a
2 GiB child cap, and no parallel build. The target's peak/final observed
footprint was 1.9 GiB; host availability was approximately 14 GiB with 132
GiB disk free and swap exhausted. These are resource observations only and
make no OOM-prevention claim. `performance_claim: none`.
