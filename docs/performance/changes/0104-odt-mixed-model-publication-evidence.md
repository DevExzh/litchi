# Change 0104: ODT mixed model-content publication evidence

Date: 2026-08-14

Status: accepted as a matched release measurement of repeated publication
versus one staged transaction for one deterministic ODT workload. The result
is not a general ODT speed, I/O, memory, cold-cache, or producer claim.

## Scope and mechanism

The selected workload applies ordinary model-backed edits and inline-content
edits to the same deterministic ODT document. The scalar control submits one
edit and commit for each plain operation, plus one publication for the complete
inline append tail. The candidate stages all operations in one transaction and
publishes one candidate. Both paths report every logical result. The medium
shape contains 80 operations and therefore measures 49 scalar publications
versus one candidate publication; the large shape contains 320 operations and
measures 193 versus one.

The timed interval is deliberately narrow. It excludes source snapshot and
operation preparation, full reopen and semantic projection, result-count
assertions, raw-member checks, durable patch/inverse/stale checks, barrier and
late-error checks, security checks, and limit gates. The raw reports retain
those case diagnostics where the harness exposes them. This evidence therefore
attributes only the repeated-publication versus one-transaction result; it does
not attribute source opening, lifecycle, serialization, compression, or
readback work.

## Matched release protocol

One clean release binary was used for all four legs: production revision
`0aa51ce526dd2010c8de7755645964755eda091c` with binary SHA-256 prefix
`4d0158283004`. The run used Linux 6.8.0-101-generic on an AMD EPYC 9575F,
CPU affinity 2, and one pinned process per leg with five warm-ups followed by
30 in-process samples. The balanced order was scalar-A, batch-A, batch-B,
scalar-B. The harness's global filesystem/cache defaults are not part of this
ODT selector's controlled protocol, and no physical cold-cache behavior is
claimed.

The four schema-1 raw reports and their SHA-256 digests are:

- [`scalar A`](../results/odt-mixed-model-scalar-a-0112.json),
  `106adbc867341a948753c000f0641fab7903aed1f1271ef24aa82b056ee84a29`;
- [`batch A`](../results/odt-mixed-model-batch-a-0112.json),
  `492ea26601f95cd951c9ecc16526c329d7e8cfe35db468547aa05638545981b5`;
- [`batch B`](../results/odt-mixed-model-batch-b-0112.json),
  `c1b4c3df6987f692964729dbd1a30e6c495e044197aea7553a74a76414d22245`;
- [`scalar B`](../results/odt-mixed-model-scalar-b-0112.json),
  `a3af20fe9f1ce05168742b4900adda30609adf314d8e8f79e154323a957ba96e`.

The machine-readable extraction is [`the compact summary`](../results/odt-mixed-model-publication-0112-summary.json).

## Results

All values below are p50 elapsed nanoseconds from the raw report. Ratios and
reductions compare each scalar leg with the batch leg in the same A/B position.

| Shape | Scalar A | Batch A | A ratio / reduction | Scalar B | Batch B | B ratio / reduction |
|---|---:|---:|---:|---:|---:|---:|
| Medium, 80 operations | 25,639,981 ns | 802,666 ns | 31.9435x / 96.8695% | 25,051,747 ns | 784,500 ns | 31.9334x / 96.8685% |
| Large, 320 operations | 2,759,243,310 ns | 21,276,074 ns | 129.6876x / 99.2289% | 2,755,886,406 ns | 20,998,049 ns | 131.2449x / 99.2381% |

The medium scalar/batch paths report 80 results each; the large paths report
320 each. Within each shape, every sample reports the same output and logical
hash for both implementations. The medium output hash is
`96ec43718d664db11b490acd72f5325801febe9f588344d72fd85e8f0016f156` and the
logical-result hash is
`ad43f6bdc31e3dca19a6bc02673add1b74be423c8a71f0f5c0e63977f3e55943`. The
large output hash is
`3d1456892d9cb851f28ec7d61a55f4024a491ab7689be98b9c300d05c9dd0427` and the
logical-result hash is
`3473568243b5194b60448e8f86c5463ab67335d26fa010be7dba63b4cacbe4a0`.

## Claim boundary

This is accepted evidence that, for this deterministic mixed model-content
workload and this measured timing boundary, one staged publication can avoid
the control's repeated publication work while preserving the reported logical
and output hashes. It is not evidence of a general ODT speedup or of reduced
source reads, decompressed/recompressed/copied/serialized bytes, allocations,
RSS, peak memory, physical device I/O, cold-cache behavior, high-latency
behavior, or real-producer interoperability. Source open, operation
preparation, reopen, lifecycle, security, limit, and durable patch gates remain
correctness boundaries outside the timed interval.
