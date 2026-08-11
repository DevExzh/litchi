# Change 0037: source-backed OPC one-Part publication

Date: 2026-08-11

Production base: `89a88883088239197e92a6dcc81465acdd4916f9`

Scope: the low-level OPC/ZIP boundary used by OOXML packages. OLE2, RTF and
ODF production code are unchanged, and iWork/IWA crates were excluded.

## Problem and change

`SourceBackedPackage` already opened OPC from an immutable positional source
without inflating ordinary Parts, but publishing one same-topology replacement
required `into_opc_package`. That conversion inflated all four ordinary Parts
in the fixed corpus, and the general writer then Deflated all four again.

The new consuming `write_part_overlay_to_stream` operation is deliberately
narrow: it replaces one existing ordinary Part while keeping its URI, content
type, relationships and package topology fixed. It validates and materializes
the selected original payload, audits an XML replacement when applicable,
regenerates only that ZIP member, and raw-copies every other local span and
central record from the already located positional archive. A small Soapberry
bridge borrows the preservation index from `IndexedArchive`, so no second EOCD
search or ordinal-to-entry guess is required; exact raw central-directory name
bytes select the member.

This is not a general mutable source-backed package. Exact payload no-ops copy
the complete source byte for byte. Real changes to signed packages and changes
on unsupported physical layouts are refused before output. Callers that need
topology, relationship or content-type changes must explicitly choose the
existing owning rewrite path.

## Publication and failure boundaries

- Replacement and adjusted Part/archive totals are checked against the same
  finite read limits before sink access.
- The selected original member is fully decoded and CRC/framing checked before
  comparison or publication. Unselected ordinary payloads are never decoded.
- Existing and replacement XML are audited before output when the selected
  content type is XML.
- Digital-signature relationships, signature paths and signature content
  types cause a typed refusal for real changes; an exact no-op remains an exact
  source copy.
- ZIP64 projection, prefixed/unsupported preservation layouts and a
  noncanonical target member cause a typed zero-output refusal. There is no
  silent full-materialization fallback.
- Source version is checked before and after raw reads and before sink writes
  throughout publication. A change after any accepted byte is reported as
  `IncompleteOutput` containing `SourceChanged`.
- Sequential writes are capped at 64 KiB. Sink failures after accepted bytes
  retain exact partial-output accounting.

The changed path retains one selected logical `Vec` and the existing validated
generated-member/compressor buffer. It does not claim zero materializations or
zero-copy output: a sequential artifact still requires reading and writing the
complete physical archive.

## Matched latency and I/O evidence

The fixed corpus contains four ordinary 4 MiB incompressible Parts and six ZIP
members in 16,783,632 source bytes. Its source SHA-256 is
`a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6`;
the middle target payload SHA-256 is
`3dbf6225021a99c1da8750a738bde21f57591c0be1a60aa510966c47ee25b098`.
Both binaries emit the identical changed archive SHA-256
`f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009`.

The primary CPU-2 ABBA run used 20 warmups and 200 samples per leg in
before-A, after-A, after-B, before-B order. The table pools 400 raw samples per
state.

| Source-backed one-Part save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 223.602 ms | 60.112 ms | **-73.12%** |
| mean | 229.034 ms | 60.512 ms | **-73.58%** |
| p95 | 270.278 ms | 62.575 ms | **-76.85%** |
| p99 | 285.578 ms | 67.454 ms | **-76.38%** |

The approximate independent-sample 95% interval for the mean delta is
`[-74.24%, -72.91%]` of the before mean. Both individual before p50 values are
223.46-223.63 ms and both after values are 60.11-60.13 ms.

Per-sample semantic materializations fall from four Parts to one. Physical
source overlap remains 516 calls / 16,782,356 bytes because the unchanged
compressed spans must still be copied. Total positional reads move only from
545 / 16,783,548 bytes to 549 / 16,785,201 bytes; no physical input-byte
reduction is claimed. Sink output is unchanged at 16,783,632 bytes, while the
bounded adapter changes 557 writes at up to 32 KiB into 461 writes at up to
64 KiB.

Raw reports are [`before A`](../results/abba-opc-source-overlay-before-a.json),
[`after A`](../results/abba-opc-source-overlay-after-a.json),
[`after B`](../results/abba-opc-source-overlay-after-b.json), and
[`before B`](../results/abba-opc-source-overlay-before-b.json). Binary and
evidence digests are indexed by
[`opc-source-overlay-sha256.txt`](../results/opc-source-overlay-sha256.txt).

## CPU and memory attribution

Matched process-wide `perf stat` legs used 20 warmups and 50 measured samples.
They include deterministic corpus construction and post-timing verification,
so their reductions are smaller than the operation-only timer but point in the
same direction.

| Counter, pooled A+B | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 37,759.45 ms | 14,216.40 ms | **-62.35%** |
| cycles | 185,440,187,247 | 69,439,717,899 | **-62.55%** |
| instructions | 473,716,788,712 | 163,788,789,730 | **-65.42%** |
| branches | 81,549,835,660 | 25,585,504,483 | **-68.63%** |
| branch misses | 1,803,157,905 | 492,656,317 | **-72.68%** |
| cache references | 12,627,900,988 | 4,737,633,250 | **-62.48%** |
| cache misses | 200,283,037 | 133,677,246 | **-33.26%** |

The sampled before profile attributes 54.27% self cycles to
`deflate_medium` and 17.65% to `longest_match`, principally below the general
all-Part `PackageWriter`. The after profile retains one compressor invocation;
its relative shares are 40.11% and 12.41%, while total captured samples fall
from 9,975 to 3,995. Both profiles report zero lost samples; restricted kernel
symbols are disclosed in the raw reports.

Matched one-shot Heaptrack reduces allocation calls from 1,497 to 1,401
(-6.41%), temporary allocations from 206 to 203, peak heap from 130.83 to
126.64 MiB (-3.20%), and Heaptrack RSS from 122.77 to 119.07 MiB (-3.01%).
Uninstrumented GNU Time ABBA maximum RSS is 120,700/120,828 KiB before and
116,884/112,460 KiB after; maximum-to-maximum is -3.26%. The retained source,
expected archive, sequential output and one required compressor buffer explain
why peak memory falls much less than CPU time.

Raw attribution is in `opc-source-overlay-*-perf-report.txt`,
`opc-source-overlay-perf-stat-*.csv`, `opc-source-overlay-*-heaptrack.txt`, and
`opc-source-overlay-time-*.txt` under `results/`.

## Correctness verification and remaining scope

- focused source-backed tests cover exact no-op identity, raw preservation of
  all unselected members including an unknown non-Part, complete OPC reopen,
  invalid XML, signatures, replacement limits, prefixed-layout refusal,
  stale sources before/during output, bounded writes and partial sinks;
- Soapberry tests cover exact UTF-8/non-UTF-8 raw names, interleaved directory
  entries, and central/local order divergence without ordinal ID mapping;
- the deterministic harness reopens and compares every Part and checks the
  exact source, target and output hashes plus all source/sink counters;
- warning-denied all-target/all-feature Clippy passes for both production
  crates and the standalone harness; and
- CI now runs one-sample smoke and 15-sample release cases with exact corpus,
  output, one-materialization and bounded-sink assertions.

This tranche does not expose a DOCX/PPTX/XLSX semantic edit facade, support
topology changes, mutate signed packages, or optimize the unavoidable raw-copy
input/output bytes. Broader source-backed CRUD, real-producer/media matrices,
atomic filesystem publication, hierarchical cache-budget charging and explicit
signature strip/resign policies remain separate work.
