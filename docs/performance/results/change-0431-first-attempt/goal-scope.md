# Change 0431 first-attempt goal scope

**Disposition: historical first candidate `e68`; superseded by the main 64 KiB
refinement now running.** This bundle's first after capture completed all 480
retained samples and passed its frozen report checks, but the delayed-range
lane regressed by 8.8–10% and resource review remains open. The global goal
remains open; no optimization or final candidate claim is accepted from this
snapshot.

## Scope and current evidence

The frozen protocol has eight serial lanes: plain and media-rich corpora each
through bytes, recently staged warm files, a 4,096-byte caller-range cap, and
a 65,536-byte caller-range cap with a 200 µs caller delay. R1 and R2 reverse
lane order. Each role uses 16 fresh processes, three warmups, and 30 retained
samples per process. The unchanged before executable has 480 samples that
passed the frozen report verifier; preliminary R1/R2 API medians differ by
less than 3.5% in every lane. The historical `e68` after executable also has
480 samples passing the frozen report verifier. Its media-rich delayed-range
p50 increased from 654.019/660.341 ms to 719.306/718.339 ms (R1/R2), an
8.8–10% regression. Plain lanes stayed within -0.005% to +1.8%; the bytes
after repeat also carries a 13.45% latency review flag and endpoint-RSS
baseline variation. These are review findings, not accepted causal claims.
The main bundle's bounded 64 KiB capture refinement supersedes this result.

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

Implementation receipts currently record 1,754 tests passing with five
ignored before the final boxed-token layout correction, plus the affected
strict, rustdoc, workspace, and formatting checks. The final consumer rerun passed
1,305 tests with three ignored. These checks establish implementation coverage;
the historical candidate performance result is recorded above.

## Historical acceptance boundary

The historical after reports were matched to the retained before executable and
corpus, preserve all source/destination/version and output gates, and retain
API timings, managed-resource journals, logical source-read histograms,
compressed and decoded transfer accounting, and portable replay. Output hashes
may differ for the newly wrapped members, but semantic output, dependency
closure, media payloads, relationships, and every untouched destination member
must remain equivalent.

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

## Remaining non-iWork work

0431 targets the Deflate-heavy media publication path identified by 0430, but
it covers one source-backed PPTX/OPC workflow. The full goal still lacks
representative CRUD completion, broader native producer/size coverage, true
cold and remote/range I/O, bounded semantic streaming and append, complete
allocator/physical-copy attribution, concurrent scaling, and remaining strict
gates. The warm-file and caller-delay lanes do not close those gaps. No claim
should be promoted from this historical snapshot. The main 64 KiB refinement
must complete its independent report, output, resource, and mutation review
before any candidate result is considered.
