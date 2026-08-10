# CFB validated lookup and bounded sector buffers

Status: accepted after measurement
Production base: `2665d572b78f0b3efd9ecfc4bd1fda09f8786ae3`

## Mechanism

`OleFile` now retains the exact `DirectoryNameData` produced by directory-tree
validation in a SID-aligned private vector. Child lookup descends the already
validated MS-CFB sibling tree and compares those cached keys. It no longer
performs a full DFS and no longer rebuilds comparison keys at every visited
node.

The same change set removes transient sector allocations from CFB open:

- FAT, DIFAT, and MiniFAT reads reuse one bounded 4096-byte stack buffer;
- MiniFAT entries are decoded directly into their final `Vec<u32>`;
- directory sectors are batched directly into the final zero-filled buffer;
- truncated final-sector zero-tail and error-order behavior are preserved.

An intermediate ordered lookup that rebuilt keys was rejected: it improved the
2,048-sibling case but made 256-sibling lookup roughly three times slower. The
cached-key implementation below is the retained version.

## Before/after protocol

Matched release binaries used identical harness sources and were run in
before/after/after/before order. Each cell contains 1,500 samples after 30
warm-ups per replicate. Corpus hashes match. Raw reports are the eight
`results/abba-cfb-final-*.json` files.

| CFB operation | Corpus | Before p50 | After p50 | Change |
|---|---|---:|---:|---:|
| open | 256 MiniFAT streams, compressible | 141.1 us | 137.3 us | -2.66% |
| open | 256 MiniFAT streams, incompressible | 141.5 us | 136.8 us | -3.30% |
| open | 2,048 root streams, compressible | 962.0 us | 974.9 us | +1.34% |
| open | 2,048 root streams, incompressible | 963.1 us | 954.1 us | -0.94% |
| read final root stream | 256 streams, compressible | 1.422 us | 0.486 us | -65.85% |
| read final root stream | 256 streams, incompressible | 1.067 us | 0.471 us | -55.88% |
| read final root stream | 2,048 streams, compressible | 7.596 us | 0.456 us | -94.00% |
| read final root stream | 2,048 streams, incompressible | 7.460 us | 0.451 us | -93.96% |

Heaptrack used 500 compressible `cfb_open` iterations per process:

| Corpus | Allocation calls | Temporary allocations | Peak heap | Profiler RSS |
|---|---:|---:|---:|---:|
| 256 streams, before | 581,921 | 184,397 | 1.35 MB | 7.60 MB |
| 256 streams, after | 530,822 (-8.8%) | 133,296 (-27.7%) | 1.35 MB | 7.66 MB |
| 2,048 streams, before | 4,410,445 | 1,301,627 | 2.66 MB | 8.29 MB |
| 2,048 streams, after | 4,141,912 (-6.1%) | 1,033,092 (-20.6%) | 2.70 MB (+1.5%) | 8.92 MB (+7.6%) |

The wide-root profiler RSS increase exceeds 5%, while peak tracked heap moves
only 1.5%. The retained name keys are the deliberate memory-for-CPU tradeoff
that prevents the rejected small-tree regression and preserves the exact
validated CFB comparison rule. No broader memory reduction is claimed.

## Correctness and limits

- 104 CFB unit tests, the complete legacy corpus, and 12 doctests pass.
- Warning-denied all-target/all-feature CFB Clippy passes.
- New tests cover cache/SID alignment, missing-cache refusal, non-contiguous
  batched reads, and truncated sector tails.
- No public CFB type, unsafe code, dependency, global state, or lock was added.

CFB still owns a mutable `Read + Seek` cursor, eagerly validates allocation
topology, materializes stream payloads into `Vec<u8>`, and requires seekable
output. Positional `ReadAt`, concurrent stream reads, and forward-only writing
remain architectural work.
