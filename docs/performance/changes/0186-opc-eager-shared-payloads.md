# Change 0186: shared eager OPC part payloads

Date: 2026-08-18

## Decision

Retain one immutable decompressed payload allocation from ordinary ZIP ingress
through OPC part construction. Before this change, serial eager OPC opening
used `LazyArchiveReader::read_shared`, cloned the cached `Arc<Vec<u8>>` into a
new `Vec<u8>`, and then wrapped that copy in another `Arc<Vec<u8>>` when it
built the XML or binary Part. The retained handoff is now
`Arc<Vec<u8>> -> Arc<Vec<u8>>`.

`SerializedPart::blob` is consequently an `Arc<Vec<u8>>`; this low-level raw
reader API is allowed to evolve on this branch. Existing `Vec<u8>` constructors
remain compatibility adapters. `XmlPart::new_shared` and
`PartFactory::load_shared` adopt an existing immutable allocation, matching
the already shared `BlobPart` representation. The explicit eager `OpenSession`
still owns each decompressed `Vec<u8>` and wraps it once without copying.

The change does not make ordinary OPC open lazy. It still validates the same
ZIP/catalog/relationship topology and decompresses every admitted Part. CRC,
size and aggregate limits, cancellation, source authorization, signatures,
unknown-member classification, exact no-op publication, and save behavior are
unchanged.

## Verification

The integrated gates passed:

- 198 OPC unit tests plus all OPC integration and documentation tests;
- 218 Soapberry ZIP library tests in independent review;
- 871 DOCX, 520 PPTX, and 769 XLSX library tests;
- strict OPC all-target Clippy with warnings denied;
- formatting and diff checks;
- independent current-tree review: SAFE.

New tests prove that ordinary eager ingress retains the archive reader's exact
payload allocation and that both XML and binary factory construction preserve
`Arc::ptr_eq` identity. The existing OPC suite retains malformed ZIP, limits,
session/cancellation, source authorization, exact save, mutation isolation,
and publication coverage.

## Measurement contract

The clean control revision is
`938146d13e537217d9473dae967a9e4cf0391b91`, release binary SHA-256
`dc26f8244208b4c348851833c151a569e850eeb22e370f5b2cfcebff0dd8fb93`.
The clean candidate revision is
`1705990aa61b0a4498b0dbaeb6ab6f7d6c5288f9`, release binary SHA-256
`9a60798138fe184c91a617a899c0f6e4d38d1a03628572d7e752db4895ddddec`.

Fresh CPU-2-pinned A1/B1/B2/A2 processes use 20 warmups and 500 retained
samples for the existing `opc_open` and `opc_open_owned` selectors on the
many-small and few-large incompressible corpora. The predeclared
p50/mean/p95/p99 same-implementation drift ceilings are 5%/5%/10%/15%.
A statistic is accepted only when both paired directions are lower and both
same-implementation drifts pass.

Deleting only revision/dirty-state and elapsed vectors produces canonical
projection SHA-256
`09ada8b6d30846d5494a1628f63f404aebabf3ae955d73a41b15f7aacbf1133d`
for all four legs with
`jq 'del(.environment.git_revision,.environment.git_worktree_dirty) | .results |= map(del(.elapsed_ns))'`.
Corpus names, archive sizes and hashes, Part counts,
configuration, compiler, allocator, machine, and affinity therefore match.

## Result

The deterministic mechanism removes one complete decompressed-payload copy per
admitted Part. A whole-process Heaptrack diagnostic over the four-Part,
16,777,216-byte few-large owned-open case records peak heap changing from
71.72M to 55.02M, a displayed 16.70M reduction. Allocation calls change
763 -> 755 and temporary allocations remain 115. The profile includes startup
and deterministic corpus construction, so its peak-RSS values are descriptive
only; the exact allocation-identity tests are the authority for the ownership
mechanism.

Few-large p50 reductions are large in both pairs: 43.42%/48.10% for borrowed
`opc_open` and 40.44%/46.98% for `opc_open_owned`. Those p50 results are
withheld because control drift is 8.87% and 5.78%, above the 5% ceiling. The
only accepted latency statistic is few-large owned-open p99, 32.99% and 43.51%
lower with 8.80%/8.27% control/candidate drift inside the 15% gate.

Many-small latency is withheld: paired directions disagree or regress and
several tail/drift gates fail. This is consistent with trading one small
payload copy for shared-ownership bookkeeping; no many-small speedup is
claimed.

No decompression, selective-open laziness, general allocation/RSS,
physical-I/O, cold-cache, throughput, scaling, real-producer, broad OOXML CRUD,
or iWork claim is made. The next higher-ROI OPC tranche remains routing
ordinary selective facade opens through deferred/source-backed catalogs rather
than inflating every admitted Part.

Artifacts:

- [summary](../results/opc-eager-shared-0194-summary.json)
- [manifest](../results/opc-eager-shared-0194-manifest.json)
- compressed raw A1/B1/B2/A2 reports and diagnostic Heaptrack/profile files
  listed in the manifest
