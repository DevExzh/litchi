# Change 0274: DOC owner public-phases ABBA hypothesis rejected

Date: 2026-08-25

Status: rejected optimization hypothesis; candidate reverted;
`performance_claim: none`

## Evidence package and protocol

The clean ABBA package is
[`results/0274-doc-owner-public-phases-abba-20260825/`](../results/0274-doc-owner-public-phases-abba-20260825/).
Its canonical summary SHA-256 is
`b1073bf107e844b3b6d4124a3f01eb10375586e0d728eca0cff7a632d72c79cf`.
The retained deterministic corpus catalog is
[`corpus-manifest.json`](../results/0274-doc-owner-public-phases-abba-20260825/corpus-manifest.json),
with SHA-256
`cc2b213e68ef7d020bb6a05a880e43c67d558344efd088f4883110476806f755`.
It is `manifest_version: 2` and catalogs the three fixed corpora
`doc-tiny`, `doc-large`, and `doc-payload-heavy`. The catalog is
Git-retained alongside the ABBA package, but is intentionally not listed
inside the four-leg package manifest; it is corpus provenance, not an
additional ABBA leg.
Control revision `c03f75b89d1e935aa87eef3c88a1ed292c05de1c` uses release
binary SHA-256
`a5c3f59ee42188aeffb1e05f7f7e68891f20000c31a8bb92a2f3342330b154df`.
Candidate revision `c2715d694a090f312391cef63142e609f7fa2249` uses release
binary SHA-256
`a91d638df0926d5a72126621073947cf8dd8f8c5dbe3be8fb746946de7fbf43b`.

The run used CPU 2, 20 warmups, 500 samples, and A1/B1/B2/A2 ordering over
the tiny, large, and payload-heavy shapes. Drift ceilings were 5% for mean
and p50, 10% for p95, and 15% for p99. The evidence package is retained to
document the decision, not as a production performance result.

## Hypothesis and disposition

The candidate tested whether removing a public-reader `Vec` clone from the
DOC owner public phases would produce a representative end-to-end benefit.
It did not. The phase vectors show that observed time shifted into source
retention, but this capture has no allocator instrumentation and does not
establish the mechanism. The local clone removal was not a sufficient
production optimization.

Only the large-shape lifecycle p50 passed the paired acceptance rule:
`+3.887%` in A1-to-B1 and `+3.253%` in A2-to-B2. The tiny-shape p50 was
adverse in both directions (`-1.035%` and `-3.469%`). Payload-heavy
directions disagree, and the means and tails are rejected for noise,
directional disagreement, or drift. The payload-heavy `open_retain`
means were about 3.65x and 3.73x the paired control means. This is descriptive
phase attribution only, not an allocator or production-performance claim.

No allocation, RSS, physical-I/O, decompression, copy, or other resource
evidence was collected for acceptance. The candidate was reverted by
`e8c0d256e` under the GOAL keep/revert rule, so the current production path
retains no optimization from this hypothesis.

## Boundaries

This change adds no selector and does not change the selectable or default
case counts: the current matrix remains 393 names and the default remains
36 cases / 198 records. No claim-registry entry or historical classification
is updated. The result does not support a broad DOC claim, an allocator/RSS
claim, physical-I/O behavior, or a production latency claim.

The retained corpus catalog binds the three deterministic generator identities
and hashes to the reports, but it contains no source files or member
inventories. Reproducing package bytes therefore requires the benchmark
generator at the pinned revisions; the evidence directory is not a standalone
source-package archive.
