# 0772 — current OPC publication baseline; ordinary-stack attribution rejected

Status: descriptive evidence only. `performance_claim: none`.
Source `6074e10e57`; [packet](results/change-0772/README.md).
No production or harness source changed. OLE2/OOXML remain the priority;
iWork is excluded.

Three serial processes, pinned to CPU 12, each measured nine samples after
two warmups for both existing OPC publication selectors and all eight generated
corpora. The unchanged harness was freshly built with release/offline/locked
arguments. Its source census, compiler, binary identity, commands and raw
reports are retained. This is not a comparison with the historical 0756 sweep.

Median of the three process p50s, in milliseconds:

| Corpus | Exact no-op publication | One-byte mutation publication |
| --- | ---: | ---: |
| tiny, compressible | 0.000050 | 0.008480 |
| tiny, incompressible | 0.000040 | 0.018760 |
| many-small, compressible | 0.000440 | 0.128471 |
| many-small, incompressible | 0.002990 | 0.149280 |
| few-large, compressible | 0.000810 | 0.747833 |
| few-large, incompressible | 2.356360 | 59.403451 |
| wide-root, compressible | 0.005750 | 1.211555 |
| wide-root, incompressible | 0.006520 | 1.221366 |

The large case has four 4 MiB parts. Mutation changes the first byte of one
part. Its process p50s are 59.353–59.445 ms. Two no-op controls exceed the
preexisting 5% spread flag: tiny/incompressible (25%, at tens of nanoseconds)
and many-small/incompressible (5.137%). Every process and flag is in
`analysis.json`; these short controls must not support precise relative claims.

## What the measurement does and does not prove

The clock covers publication into a pre-reserved bounded non-seek sink. Package
open, mutation, expected-output construction, output equality checks and sink
inspection are outside it. No filesystem sync, cold-storage behavior, allocation
lane, RSS, concurrency scaling or native Office interoperability was measured.
The selector checks output against its own writer's expected bytes and stable
sink summaries. This revision does not expose the published digest or an
independent untouched-member oracle; it is insufficient for accepting a new
preservation optimization without strengthening those checks.

Source review corrected an initially suggested optimization: lazy decoding of
the old target happens before timing (`get_part_mut` forces it, and expected
serialization also precedes the loop). Timed equality checks see cached bytes
and can stop at the first differing byte. Removing that decode cannot explain
or improve the reported repeated-publication time. Equal-byte replacements must
still retain their source framing; replacement provenance alone cannot justify
skipping equality. See `source-review.md`.

## Rejected attribution

One ordinary-binary `cycles:u` profile used 199 Hz, DWARF stack snapshots of
32,768 bytes, 30 samples and no warmups. Disassembly identifies the exact timed
writer call and its return offset in `run_opc_mutated_save`; the untimed
expected serialization is a different call.

Of 584 sampled events, **none** recovered that timed caller. Two recovered a
different root offset in the untimed comparison; 582 lacked a recognized root.
All 10,567,307,723 sampled period units are accounted for, but the profile is
rejected for attribution. A missing root does not mean a sample was outside
timing. No Deflate fraction, removable-cost estimate or causal speedup follows.

The first offline decoder used an unsupported `--sym-offset` option. Its failure
is retained; the same raw capture was decoded successfully with the supported
`symoff` field, without rerunning the workload. Raw data is retained losslessly
as gzip, with encoded and decoded hashes.

No candidate was accepted. Current evidence prioritizes finishing the already
implemented, owner-authorized durability policy integration over another
descriptive profile attempt. Any renewed optimization of this selector needs a
qualified caller trace and stronger preservation evidence first.

## Validation and cleanup

The analyzer replays all three receipts, 48 case results and 432 samples,
including binary identity, CPU affinity, source revision, serial schedule,
corpus/sink stability and independently recomputed p50/p95/mean statistics.
The stack qualifier accounts for every event and records rejection explicitly.
This is evidence validation, not a new production-test result.

The owned build target was removed after rechecking the executable's hash and
size: 1,244,751,311 logical file bytes. The profiler disabled build-ID cache
creation. One unused agent-owned oracle draft was removed; it was not integrated.
The durability worktree and all unrelated files were preserved.
