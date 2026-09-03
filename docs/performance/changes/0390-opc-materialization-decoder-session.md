# Change 0390: OPC materialization decoder session

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope and mechanism

Unmanaged source-backed OPC full materialization now creates one opaque,
operation-scoped indexed-read session after managed-package refusal. The
session lazily owns one Deflate decoder, resets it for each bounded deflated
entry, and is threaded only through cold loader/bypass reads. Stored entries
bypass the decoder. Cache hits and waiters do not use the session. One-shot
indexed reads retain their existing fresh-session behavior.

The session preserves lookup order, limits, fallible output reservation,
Store/Deflate dispatch, CRC and declared-size verification, accepted-size and
error precedence, source accounting, cancellation, cache publication and
rollback, and final freshness checks. Managed reservation-bearing handles
still refuse escape before ordinary payload reads. No broad/default OPC or
iWork path is changed.

## Matched allocator evidence

The allocator target binds control and candidate to baseline revision
`2c0fd89c7c0e873ada1e62d58ac454c59c83b8bf`; the candidate adds the exact patch
SHA-256
`d825426486c50788c71dd8bb7ef76045642448216f8a8136a059556639a4e4b4` over the
two implementation files. Control blobs are
`crates/litchi-opc/src/source_backed.rs`
`2b1058c42eee28c5c9eaeb7afe9ca81702425676` and
`crates/soapberry-zip/src/office.rs`
`f0a2f6b6f7351356c694b5dd184729e40b50d917`; candidate blobs are
`d9b2971b2b7f8de7c1ad53196f7bbb3ff2b13aeb` and
`33115a179546257e2292a78300a4382a86a83bcc`. Both release binaries use
`rustc 1.98.1 (48a229cea 2026-09-01)` and the operation-scoped system
allocator. Their SHA-256 values are control
`1f4c917cc16e33b10e145e70223d6325b02a8f395a8699bfdb395155c67a2544` and
candidate
`50543d9c5aa3585f26c8d6f3feb2725b6e0a0dcde9957a1b8e04f464b618c9a7`.

Each corpus used three warmups and 15 retained samples. The measured avoided
decoder is exactly `-2` allocation calls and `-80,320` allocated bytes per
avoided decoder. The control-to-candidate vectors are:

| Corpus | Parts | Allocation calls | Allocated bytes | Removed calls / bytes | Logical reads / returned bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| tiny compressible | 3 | 47 -> 43 | 245,100 -> 84,460 | 4 / 160,640 | 15 / 318 |
| many-small incompressible | 256 | 3,097 -> 2,587 | 21,017,729 -> 536,129 | 510 / 20,481,600 | 1,280 / 276,992 |
| few-large incompressible | 4 | 61 -> 55 | 17,102,213 -> 16,861,253 | 6 / 240,960 | 532 / 16,782,540 |

Logical source reads, returned bytes, and materialized Part counts are
invariant across all retained samples. The [compact summary](../results/opc-source-materialization-decoder-session-0390-summary.json)
binds source, binary, corpus, vector, and artifact hashes. The
[compressed reports and catalogs](../results/change-0390/) retain only these
allocator-enabled report pairs and their corpus sidecars. The [checked
comparison](../results/opc-source-materialization-decoder-session-0390-comparison.json)
uses the existing operation allocator policy: three matched results, 15
metrics, and zero regressions.

The in-memory rederivation is independently checked with:

```sh
python3 -m unittest \
  tools.test_perf_compare.PerfCompareTests.test_checked_0390_decoder_session_allocator_evidence_rederives_in_memory
```

## Correctness and claim boundary

Reusable-session tests cover interleaved Store/Deflate reads, data-descriptor
entries, verifier failures, corrupt and truncated Deflate streams, and declared
size overrun/underrun recovery. The existing one-shot indexed-read suite still
covers limits, invalid IDs, injected short reads, and allocation failures
through the fresh-session compatibility adapter. Source-backed OPC cache-hit,
loader, bypass, failure, freshness, cancellation, managed no-read, and Arc/COW
behavior remains covered by the existing focused tests. The checked comparator
test rederives the comparison from all six compressed report members in memory;
no decompressed report is checked in.

Latency and throughput, operation-local peak/RSS, copied bytes, decompressed
bytes, physical I/O, and broad/default OPC conclusions are withheld. The
allocator vectors are mechanism evidence only and do not authorize an
end-to-end performance claim.
