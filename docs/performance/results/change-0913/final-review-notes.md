# 0913 pre-capture final-review notes

Status: **conditional pass; execution gates remain open**. This is an
independent static review of the candidate, retained custody and current
measurement helpers. It ran no build, test, benchmark workload, profiler, or
native capture.

The candidate is the bounded change from parent
`66b944b07fcf4c6fa37f10e7ed83a02417455f57` in
`crates/soapberry-zip/src/office.rs` and `crates/litchi-opc/src/payload.rs`.
The owned route calls the specialized `IndexedArchive<Arc<Vec<u8>>>` reader;
the generic `ReaderAt` constructor still selects the existing `MemberSpan`
policy. The direct reader reuses the indexed local-header and ZIP64 parsing,
source-bound checks, output reservation, Store/Deflate decoding, descriptor
handling, CRC/size verification, accounting, and final-size checks. The
static source review found no unsafe code, mutable shared decoder, new cache,
source-publication change, or ordinary format API expansion.

The frozen custody currently binds the 9,358-file candidate inventory to the
9355-file 0910 baseline plus the four declared overlays and marker. The plan
is frozen before capture with 48 serial children: 12 baseline/candidate pairs
for each of `materialized` and `owned`, six AB and six BA pairs, 30 warmups,
20 samples, and CPU 12. The candidate build receipt retains the candidate
binary and binds the unchanged 0910 control binary. The quality-v8 receipt is
reported as pass with 8,096 passed tests across 330 suites and 50 ignored
tests; source recovery reports 9,358 candidate files, 1,445 temporary quality
fixtures admitted and removed, and no production-file change. These are
retained receipts reviewed for custody, not checks rerun by this review.

The current report and review bindings are:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `docs/performance/0913-pptx-owned-direct-read.md` | 6,787 | `ccb192b06ac3eed4534bc74181eccbe85364a1e1e8d7ba5ca7290260b55df91d` |
| `docs/performance/results/change-0913/source-review.md` | 7,536 | `49c7151ec5f5d76438b3e71fb7c3d4e72c1077dca9ead6f2e539c70ea0e6cdbc` |
| `docs/performance/results/change-0913/protocol-review.md` | 11,490 | `5363497f204967fc3f831903fa0af557cca0af0090a89948c0b36c84b139e126` |
| `docs/performance/results/change-0913/capture-plan.json` | 10,141 | `56ab8a214388fdc8686e7ad3fbc0077a36e23a1213a29ba99dc6df8c003c054f` |
| `docs/performance/results/change-0913/source-v1.json` | 1,758,278 | `68de7c9416ba7763606a64780bde4e3de8476b02a809234b4a36975b9612af82` |

The following gates are still required before a final result review can
admit a timing or speedup claim:

1. Run the declared `capture.py` entry point and retain all 48 successful
   child receipts. Each child must bind the frozen source, route, binary,
   plan, explicit environment, source/output/sink/semantic guards, process
   metrics, and artifact files. A failed or missing child rejects the packet.
2. Produce `analysis.json` from the retained receipts and review every
   route-specific p50/p95/p99 and whole-process maximum-RSS regression flag.
   The owned adoption gate is baseline-over-candidate total median at least
   1.05 with a 95% paired bootstrap lower bound above 1.0. The materialized
   and owned routes remain separate controls; no cross-route claim is allowed.
3. Run `independent-stats.py` over the raw receipts and require its output to
   agree byte-for-byte with the retained independent result. The independent
   check must retain the 12 paired medians per route, seed 913, 10,000 draws,
   source/output/sink guards, and whole-process RSS values.
4. Bind the final report, primary analysis, independent statistics, both
   executable descriptors, source manifest, plan, child receipts and raw
   artifacts in `result-review.json`. The report must describe only this warm
   generated fixture and shared host, with no allocation, native-cycle,
   cold-cache, producer-diversity, cross-route, or general speedup claim.
5. Run fail-closed replay before cleanup, remove only marked temporary source
   and build roots after replay, then run replay again and retain matching
   before/after receipts plus the cleanup receipt. The shared iWork conflict
   state and unrelated worktrees/files remain outside this packet's cleanup.

Two pre-existing qualifications remain material to the final review. ADR
0030 says package cloning creates independently decoded part cells that may be
charged twice, while the current `PartPayload::Deferred(Arc<DeferredPayload>)`
implementation shares one `OnceLock` outcome and one aggregate charge across
clones. The candidate does not introduce this mismatch; the final review must
record the current implementation as the governing behavior or retain an ADR
resolution. Also, the supplemental adversarial test's malformed-member
failures are exposed through an outer I/O error; its family/diagnostic parity
does not by itself replace typed field assertions for each refusal class.

The separate user-authorized iWork cherry-pick and its quality receipts do not
contribute evidence to this performance candidate. No native result, adoption
decision, or copy-removal speedup is admissible while capture and replay
remain pending.
