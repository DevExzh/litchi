# Change 0175: immutable CFB atomic-save publication

Date: 2026-08-17

## Decision

Retain the immutable-owned specialization for
`ValidatedOverlayPlan::save`. Plans created from
`SharedOleFile::open_owned(Arc<[u8]>)` now omit the initial and
post-flush/pre-rename complete fingerprint scans because their source bytes
cannot change while the plan is alive. Generic positional sources retain both
fences, including the late fence against stable-token mutation.

The owned path still reopens and validates the candidate during planning,
reads the complete source in 64 KiB emission chunks, hashes both source and
target while emitting, checks exact read/write progress, flushes and fsyncs the
sibling temporary file, atomically replaces the destination, and syncs the
parent directory. Composed views keep their normal fingerprint preflight.

## Deterministic work reduction

For the fixed 16,913,408-byte CFB filesystem corpus, each effective owned
atomic save removes exactly two complete source scans:

- 33,826,816 logical source bytes;
- 34 one-MiB fingerprint reads; and
- two source/target SHA-256 fingerprint pairs.

The remaining emission proof performs 259 64-KiB-or-smaller source reads; the
former owned publication performed 293 reads. These are code-derived logical
work counts, not physical-I/O claims.

## Harness and correctness

The opt-in `cfb_file_owned_same_length_overlay_atomic_save` selector raises the
selectable matrix from 319 to 320 without changing the historical default 36
cases / 198 records. Filesystem ingress is included in total operation time
and occurs before the nested CFB open/plan/publication timers. The selector
records immutable ownership, source length/hash/version, one changed span,
published bytes, output hash, semantic CFB reopen and untouched-stream checks,
fresh-child cache state, and zero fabricated `ReadAt` counters.

Focused atomic-save tests distinguish the one-pass owned path from the
three-pass generic path. Warning/deprecation-denied CFB and OLE-common checks,
formatting, diff checks, and an independent provenance/atomicity review pass.
The reviewer confirmed that arbitrary `ReadAt` sources cannot acquire the
immutable marker and that flush/fsync/rename/parent-sync behavior is unchanged.

## Clean release A/B/B/A

The control is clean revision `80ddc3170`; the candidate is clean revision
`a92cf40b1`. Binary SHA-256 values are `ac16c311...` and `d1ac201c...`.
Every leg is pinned to CPU 2 on the AMD EPYC 9575F host, exposes one logical
CPU, uses 20 warmups and 30 fresh-child samples per warm and advisory
`cold-requested` state, and produces exact output SHA-256
`7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`.

Positive values mean lower candidate p50 latency:

| Boundary | A1 -> B1 | B2 -> A2 | Same-implementation control drift | Decision |
|---|---:|---:|---:|---|
| total, warm | 26.07% | 15.08% | 14.16% | descriptive only |
| total, cold-requested | 24.07% | 13.51% | 11.29% | descriptive only |
| atomic-publication phase, warm | 40.61% | 28.58% | 18.96% | descriptive only |
| atomic-publication phase, cold-requested | 39.72% | 29.22% | 17.28% | descriptive only |

Both candidate directions are materially lower, but the control drift exceeds
the repository's 5% gate. The deterministic work reduction is accepted; no
acceptance-grade latency claim is made. `cold-requested` is advisory cache
evidence and does not prove physical cold-device I/O.

## Withheld scope

This does not specialize generic/mutable sources, CFB construction,
topology-changing publication, allocation/RSS, physical I/O, compression,
throughput, or real-producer behavior. The timed total includes the full
filesystem read into owned bytes, while the nested phase samples exclude that
ingress. The source counter is intentionally not applicable for immutable
slice ownership.

Artifacts:

- [summary](../results/cfb-owned-atomic-save-0175-summary.json)
- [manifest](../results/cfb-owned-atomic-save-0175-manifest.json)
- raw A1/B1/B2/A2 reports listed in the manifest
