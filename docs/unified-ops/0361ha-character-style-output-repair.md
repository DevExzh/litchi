# 0361ha — Return character-style buffer growth for repair

- Base: `61515148e1`; parent proposal: [0361h](0361h-character-style-required-limits.md).
- Status: v2 allocation regression returned; v3 accepted after correctness,
  allocation and native qualification in [0361h](0361h-character-style-required-limits.md).

The first protected character-style candidate preserves all 20 paired
fixture/output files exactly. Read operations retain their allocation counts
and requested/live bytes. Its restyle writer nevertheless regresses memory:

| Fixture | Allocations/call, base → v2 | Requested bytes/call | Peak live bytes |
| --- | ---: | ---: | ---: |
| Two runs | 22 → 18 | 1,250 → 1,327 | 554 → 622 |
| 64 runs | 287 → 222 | 29,427 → 30,684 | 7,579 → 7,852 |

The common probe counts 100 operations per process, frees every returned
value inside the measured interval, and reports zero retained bytes. These
are allocator measurements with concurrent builds recorded, not latency or
process-RSS claims. The v2 candidate has no accepted native timing. Its
protected binaries, source manifests and raw allocator outputs remain under
the external `char-style-0361h` evidence directory.

The repair admits exact table and final-output sizes before reservation,
writes each table entry directly into the table buffer, and releases obsolete
run vectors before output construction. Stylesheet registration plans its
replacement entry while scanning the parent records, then reserves the
complete output once. Variation construction similarly plans all nested
lengths and writes into one admitted final buffer. Every output append stays
within the admitted size; unchanged fields retain their original bytes.

Tests exercise all 128 combinations of the seven character properties with
both a missing parent and a maximum-width identifier, exact output limits,
parent-entry length-prefix changes, and overflow refusal before reservation.
The rebuilt v3 probe again preserves all 20 paired fixture/output files.
The [retained allocation summary](results/0361h/allocation.json) includes
both candidates and the base. Two-run restyle now uses 10 allocations,
988 requested bytes and 480 peak live bytes per call; the 64-run case uses
147 allocations, 26,316 requested bytes and 4,704 peak live bytes. Stylesheet
registration reduces calls, requested bytes and peak live bytes at both
sizes. Full variation construction uses one allocation and 64 bytes, down
from 14 allocations, 304 requested bytes and 144 peak live bytes. Read-only
routes stay identical to the base. Every route retains zero bytes after its
result is dropped.

The full consumer checks pass: 57 raw unit tests, eight compile-fail doctests,
1,015 owner tests, all-target/all-feature Clippy and the raw glossary. All
11 protected native routes improve; restyle medians are 0.6374× and 0.8562×
for the small and larger fixtures. The earlier clean blocks from the
interrupted first attempt remain included. This closes the returned v2
regression without accepting a performance exception.
