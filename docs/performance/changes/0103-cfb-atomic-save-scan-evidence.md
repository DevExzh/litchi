# Change 0103: CFB atomic-save scan evidence

Date: 2026-08-14

Status: accepted as exact logical source-read evidence for one atomic-save
case. No latency, speed, allocation, RSS, peak-memory, or physical-cold result
is claimed.

## Scope and mechanism

The atomic `ValidatedOverlayPlan::save` path previously performed four complete
source/fingerprint scans of an `N`-byte CFB artifact: the initial preflight,
the source/target hash while emitting, the post-emission preflight inside the
shared writer, and the mandatory post-flush/fsync pre-rename preflight. The
follow-up keeps the initial preflight, the output-time source/target hashes,
and the final pre-rename check, while skipping only the duplicate
post-emission preflight in `save`: `4N -> 3N`. Direct `write_to` retains its
post-emission preflight and is unchanged.

The change preserves the source and target fingerprints, candidate validation,
output-time hashing, flush/fsync ordering, atomic sibling replacement, late
mutation detection, temporary-file cleanup, destination preservation on
failure, and typed partial-output behavior. The production revision is
`4ededfa245173d2e03f357b91ebf81b867ebc01f`, compared with
`32e5a9f819399111bcb2cd70b8c2b7f6887c1773`. The matched release binaries were
identified by the supplied SHA-256 prefixes `b95c5d3e` (before) and
`88a5adb7` (after); the raw reports retain the exact source revisions and
environment.

## Matched release evidence

The case is `cfb_file_same_length_overlay_atomic_save` over the deterministic
`ole-common-few-large-incompressible` CFB/OLE2 corpus: five entries, 4 MiB
entry shape, 16,913,408-byte archive, and one 36-byte same-length replacement.
The run used release binaries pinned to CPU 2 on Linux 6.8.0-101-generic,
AMD EPYC 9575F, ext2/ext3, warm cache state, five warm-ups, and 30 fresh-child
process-isolated samples per leg in the order before-A, after-A, after-B,
before-B. All four raw schema-1 reports and the compact extraction are linked
below:

- [`before A`](../results/cfb-save-atomic-scan-before-a-0112.json), SHA-256
  `8930a63782469b403c519fe98af7e2a3a677fce726bae90fb2d83f68f48ccbbf`;
- [`after A`](../results/cfb-save-atomic-scan-after-a-0112.json), SHA-256
  `40c6c552487d78a0989b7b077117351628540caef3aff99e217049468b670a7a`;
- [`after B`](../results/cfb-save-atomic-scan-after-b-0112.json), SHA-256
  `43878dd4da31d1b6734fd5f3dff0c58a27f2131912ac4e21414d50c12b64fc4e`;
- [`before B`](../results/cfb-save-atomic-scan-before-b-0112.json), SHA-256
  `7c0565cfedfeb458e885659bd3b300fed60968ba9220b3892a573a880dcccd31`;
- [`compact summary`](../results/cfb-save-atomic-scan-0112-summary.json).

The before legs each issue exactly 2,084 logical `ReadAt` calls requesting and
returning 101,751,908 bytes. The after legs each issue exactly 1,825 calls
requesting and returning 84,838,500 bytes. Therefore this matched case removes
16,913,408 logical source bytes (16.6222%) and 259 logical calls (12.4280%).
Every leg publishes exactly 16,913,408 bytes with the same output SHA-256,
`7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`.

## Claim boundary

The balanced latency directions do not agree: before-A to after-A p50 is
143,425,701 ns to 148,870,583 ns (+3.7963%), while before-B to after-B is
164,880,142 ns to 148,368,923 ns (-10.0141%). This evidence therefore makes
no latency or speedup claim. Parent-process wall time is retained in the raw
reports for context and is not a speed measurement. The process-level
`read_bytes` observations are warm-cache counters (zero in the B legs and
zero-to-4 KiB in the A legs), not physical device-I/O evidence.

No allocation, RSS, peak-memory, physical cold-cache, high-latency, storage
throughput, decompressed/recompressed/copied-byte, or general DOC/XLS/PPT
semantic CRUD claim follows from this run. Direct sequential `write_to`
behavior is not included in the optimization claim and remains covered by its
existing three-scan guard. The exact logical `ReadAt` reduction and identical
published output are the accepted results.
