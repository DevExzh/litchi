# Change 0431 goal scope

**Disposition: accepted scoped measured optimization; the global goal remains
open.** Change 0431 adds a bounded compressed source-part transfer path for
source-backed OPC/PPTX publication. The refined matched comparison passed 16
fresh processes and 480 retained samples per role with no API repeat or
regression flag above 5%. Portable verification passed before cleanup, and task scratch removal
completed. The result does not establish an allocation, RSS, physical-copy,
or zero-copy claim.

## Scope and current evidence

The frozen protocol has eight serial lanes: plain and media-rich corpora each
through bytes, recently staged warm files, a 4,096-byte caller-range cap, and
a 65,536-byte caller-range cap with a 200 µs caller delay. R1 and R2 reverse
lane order. Each role uses 16 fresh processes, three warmups, and 30 retained
samples per process. The unchanged before executable and refined candidate each
have 16 fresh processes and 480 retained samples, with the before role passing
the frozen verifier and the matched comparison passing. The candidate source
revision is `0556401e2`. API p50 values below are R1/R2 milliseconds; the
percentages are the corresponding before-to-after changes:

| Lane | Before | Refined after | Change |
| --- | ---: | ---: | ---: |
| media-rich bytes | 252.841 / 257.666 | 28.294 / 28.263 | -88.809% / -89.031% |
| media-rich warm file | 255.177 / 255.196 | 34.008 / 34.393 | -86.673% / -86.523% |
| media-rich short range | 263.248 / 254.163 | 29.694 / 29.791 | -88.720% / -88.279% |
| media-rich delayed range | 654.019 / 660.341 | 527.453 / 527.441 | -19.352% / -20.126% |

Plain lanes change by -1.045% to +0.091%. No API lane exceeds the 5% repeat
or regression review threshold. The first attempt and its 8.8–10% delayed-range
regression remain preserved in the sibling
[`change-0431-first-attempt`](../change-0431-first-attempt/) bundle; the
refined capture supersedes its result. Peak media RSS across roles is
783.2–785.3 MiB, but endpoint baselines vary, so no causal memory conclusion
is accepted.

The implementation has three bounded ownership steps:

* `soapberry-zip` issues a private Store-or-Deflate token only after bounded
  compressed-range capture, layout and descriptor checks, exact decoded-size
  and CRC validation, and Deflate `StreamEnd`/exact-consumption validation.
* `litchi-opc` authorizes the token against source lineage, revision, content
  type, decoded bytes, execution context, and managed budgets. It fences the
  token during publication and preserves typed partial-output errors.
* PPTX retains its existing decoded image/chart, XML, relationship, topology,
  and candidate checks, then attaches an authorized token for eligible media
  or leaf charts. XML and relationship additions still use the logical writer.

The transferred payload is buffered into a new canonical destination wrapper;
it is not a raw source-member copy and does not establish zero-copy ownership.
Untouched destination members retain their existing physical preservation
contract. The source-oracle audit separately checks decoded semantics and
physical untouched members; it excludes only the three regenerated package-level
XML members whose wrappers may differ.

Refined implementation receipts record 1,755 tests passing with five ignored;
the strict and formatting checks pass, and the ASAN suite records 1,000 passes.
Portable verification passed before cleanup, and task scratch removal
completed. These checks and the
matched comparison establish only the scoped API-duration result below, not a
whole-process or general memory result.

## Accepted scope and remaining verification

The refined comparison is accepted only for the named source-backed PPTX/OPC
corpus, build, machine, warm/file/caller-range lanes, and API timer boundary.
It preserves the source/destination/version and output gates and retains API
timings, managed-resource journals, logical source-read histograms, compressed
and decoded transfer accounting, and comparison reports. Output hashes may
differ for newly wrapped members, but semantic output, dependency closure,
media payloads, relationships, and every untouched destination member remain
equivalent. Portable final replay and cleanup still need to complete before
the evidence bundle is final.

The transfer must remain fail closed. Source mutation, cancellation, read or
work limits, malformed layout, CRC/size or descriptor failure, unsupported
methods, signature or active-content policy, and topology/relationship refusal
must remain typed errors. Only the explicitly scoped OPC memory-admission
refusal may choose the existing logical recompression path. If a publication
writer accepts a prefix and then fails, `IncompleteOutput { written, .. }`
must retain precedence over a final source-fence error.

The ZIP wrapper must derive local/central metadata, offsets, descriptors and
ZIP32/ZIP64 framing from the destination layout. It must not copy source local
headers, flags, timestamps, extras, or offsets. Compressed token capacity,
decoded validation storage, generated wrapper capacity, names, output framing,
and checked integer conversions must remain inside the configured hierarchical
budgets. The reservation is an explicitly modeled payload bound, not a
whole-operation heap or RSS bound.

The measured result is an API-duration result for this matched workflow. It
does not claim that all physical copies, allocator work, RSS, or whole-process
latency were reduced.

## Remaining non-iWork work

0431 targets the Deflate-heavy media publication path identified by 0430, but
it covers one source-backed PPTX/OPC workflow. The full goal still lacks
representative CRUD completion, broader native producer/size coverage, true
cold and remote/range I/O, bounded semantic streaming and append, complete
allocator/physical-copy attribution, concurrent scaling, and remaining strict
gates. The warm-file and caller-delay lanes do not close those gaps. Portable
final/cleanup verification remains the immediate evidence task; broader native,
cold/remote, CRUD, streaming, scaling, and strict-gate work remains open.
