# 0549 checked bitset test-and-mark campaign

The previous goal turn made progress: 0548 rejected the checkpoint collector
using native and instruction evidence, retained the public corruption guard,
committed and cleaned. OLE2/OOXML remain active; ODF is deferred until that goal
completes and iWork is excluded.

This experiment keeps exact visited-bit cycle detection and tests a private
combined membership/mark operation only in SectorChainScratch::collect_exact.
Assembly must demonstrate actual emitted work reduction; source duplication is
not itself evidence. Preflight, reservation order/resource labels, zero-fill,
invalid-index/first-cycle/marker ordering, scratch reset/reuse and ownership
validation remain unchanged. No new API, dependency or unsafe code is allowed.

The already tracked guard example is common to both stages. Root runs all
builds and measurement children serially, with no live Rust or frozen-driver
mutation during those processes. Each stage starts with release guard Clippy,
release guard build and all 16 one-sample oracle smokes before main builds.
The baseline completes before the reviewed candidate is applied. Source,
compiler, plan, binary and command/output hashes bind all evidence.

The full matrix and thresholds are inherited unchanged from the completed
0548 campaign: two-repeat ABBA native and guard order, 20/1000 main warmups/
samples, separate 3/30 allocation samples, 0/5 constructor profiles, and
whole-child hardware diagnostics. Four primary XLS p50 workflows must each
improve at least 3% in both repeats. XLS-owned constructor Ir and collector
self Ir in XLS-owned and CFB-few-large must fall in both repeats. Allocation
calls/bytes/incremental peak must not grow materially. Guard invalid p50 and
mean must remain within 4x same-invalid and 2x baseline-valid at each size and
repeat. Every >5% adverse and absolute repeat variation remains reviewed.

Correctness gates include exact helper/boundary and exhaustive chain oracles,
public malformed guards, and the 15 declared quality commands. Candidate
retention requires every gate. Rejection restores the exact measured baseline
runtime-plus-test source and runs all 15 final quality checks; the candidate
snapshot and failed evidence remain. No changed final rebuild may reuse old
performance captures. Final baseline equality permits the exact retained
baseline binaries to supply their already captured evidence.

No sanitizer campaign is claimed when cargo-fuzz/nightly are unavailable.
Current host tool observations record that availability. Guard timing stops
after OleFile::open, before result comparison/destruction, and has no allocator
instrumentation. Profile setup and termination dumps are excluded from timed
constructor attribution. Hardware counters/RSS are whole-child diagnostics;
there is no operation-local hardware or stable-tail/scaling claim.

All raw captures and intermediate attempts are preserved. Canonical analysis
runs use exclusive-create outputs and reproducible report replay. After final
source/quality/review/decision verification, six retained binaries are hashed,
accessible process references checked, and the sole target removed. Seal the
complete bundle and documentation, commit exact reviewed bytes, and verify
again after commit. The broad goal remains active.
