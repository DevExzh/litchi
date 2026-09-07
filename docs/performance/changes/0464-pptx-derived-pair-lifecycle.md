# 0464 — descriptive PPTX derived-pair lifecycle evidence

0464 records a harness-only PPTX source/destination pair lifecycle. It changes
no production PPTX implementation. The pair is a positive control with
explicit source and destination owners, but the destination is derived from
the same LibreOffice QA source archive; it is not an independently authored
native pair. The independent ZIP/XML/relationship oracle checks the declared
copy and output scope, but cannot establish Microsoft Office acceptance.

The frozen matrix has eight reports and 240 retained samples: R1 forward and
R2 reverse lane order, normal and operation-scoped allocator binaries, bytes
and bounded logical-range providers, three warmups and 30 samples per lane on
CPU 2 with one worker. The range adapter caps returned bytes at 256 with zero
fixed delay; it is not a network, physical-storage or cold-cache experiment.
Every formal output has three slides, 55,891 bytes, and passes the independent
package oracle.

| Lane | R1 / R2 API-sum p50 |
|---|---:|
| Normal bytes | 1.9536 / 1.9380 ms |
| Normal logical range | 1.9872 / 1.9865 ms |
| Allocator bytes | 2.0768 / 2.0705 ms |
| Allocator logical range | 2.1182 / 2.1213 ms |

The API sum is the checked sum of source open, destination open, plan and
publication calls; it is not a contiguous end-to-end timer. Input loading,
adapter construction, sink reservation, artifact writes, oracle work and
teardown are outside that sum. The normal binary intentionally exposes no
allocator totals; allocator figures are operation-scoped and are not inferred
for normal lanes. No compression or copied-byte counts are claimed.

Across all 120 allocator rows, summing the four allocator regions gives exactly
8,623,012 allocated bytes, 9,616 allocation calls and 1,455 reallocation
calls. This does not infer a lifecycle peak by summing phase peaks. Bound GNU
time resource logs report maximum RSS in kB for R1 normal-bytes/normal-range/
allocator-bytes/allocator-range as 16,912/16,908/17,068/17,152 and for R2 as
16,900/16,684/16,612/16,584. The overall 16,584–17,152 KiB (~16.2–16.75 MiB)
range is process-level and includes setup, oracle work and teardown. The
summary does not derive this GNU-time resource field; raw
`captures/R1/*/resource.log` and `captures/R2/*/resource.log` files are bound
in the bundle.

The final source review finds no harness source blocker at custody epoch
`a35f4507a74e7678f29f91dd6857d07a87579ba872785a1eaf19e296540debbe` and
records the final warning-denied harness Clippy receipt. Harness validation
records 391 passing tests with one ignored; all 240 formal samples and focused
smoke (four positive and seven negative cases) pass.

LibreOffice 26.2.5.2 saves the formal three-slide output. Native-r2 passes the
scoped application-save plus source-backed/eager validation of all three slide
counts, slide sizes and ordered text; its saved output is 39,830 bytes with
SHA-256 `d2ca3109448f5a0797f00561beff4eafa84616f7dbf897ff2456667fbf555915`.
All three post-save source-backed image inventories are unavailable because
`UnsafeEdit` / `source-backed picture inventory` refuses markup compatibility.
The failed native-roundtrip-r1 receipt and raw saved output remain preserved;
image equivalence and rendering compatibility are unproven, so this is not
full Office acceptance. The supplemental inventory diagnostic is an
inventory-only source epoch
(`28f1843728fe0a84e91c2bd5498127f3a872b46651c738ef9abb74b264320bdf` versus
the measured `a35f4507a74e7678f29f91dd6857d07a87579ba872785a1eaf19e296540debbe`);
its diagnostic build and warning-denied Clippy receipt pass. Source
compatibility records 7,031 unchanged files, and the measured copy
Rust/binary/captures are immutable with no retiming. The documentation receipt
and the owned performance-crate scoped
format check pass. A broad `cargo fmt --all --check` receipt remains failed
only at pre-existing formatting in the out-of-scope iWork
`crates/litchi-keynote/src/document.rs`; no full-workspace format pass is
claimed. Boundary checks pass. The corrected `profiling-r1` receipt supersedes
the initial UID-0-unavailable diagnostic. It passes a whole-process user-counter
and DWARF profile over three warmups and 30 retained samples (33 checked
iterations); separate counter and sampling runs add 60 diagnostic samples
outside the 240 formal samples. Totals are instructions 1,758,215,198, cycles
631,274,012, branches 327,372,796, branch misses 3,514,541, cache misses
2,295,690, and page faults 1,709. PMU counters ran at 83% because of
multiplexing, and the raw zero cache-reference value is uninterpreted. The
20-sample cycle symbol view places `sha2::sha256::x86_sha::compress` at
19.96% and `zlib_rs::inflate::inflate_fast_help_avx2` at 14.25%; the
whole-process scope includes binary hashing, setup and oracle work, so this is
diagnostic only and does not establish operation hotspots or stable ranking.
Sealed precleanup, fresh-copy portable verification and resealed-summary
tamper rejection pass. Owned cleanup removed 12 files totaling 252,885,687
bytes plus two archived audit temporaries. The batch adds no selector or
corpus coverage; the registry remains
439 selectors / 36 defaults, the full non-iWork goal remains open, and iWork is
excluded. See the [summary](../results/change-0464/summary.json),
[protocol](../results/change-0464/protocol.json),
[pair manifest](../results/change-0464/pair.json), [fixture provenance](../results/change-0464/fixture-provenance.md),
[source review](../results/change-0464/source-review.md), [native-r1 receipt](../results/change-0464/native-roundtrip/receipt.json),
[native-r2 receipt](../results/change-0464/native-roundtrip-r2/receipt.json),
[profile summary](../results/change-0464/profiling-r1/profile-summary.json),
[top symbols](../results/change-0464/profiling-r1/samples/top-symbols.txt), and
[source compatibility](../results/change-0464/source-compatibility.json).
