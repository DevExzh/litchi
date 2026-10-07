# 0957 final allocation diagnostic source review

- Reviewed at: 2026-10-07T12:49:44Z
- Source under review: tools/perf-baseline/src/xlsx_managed_compression_allocation.rs
- Previously reviewed source: revisions/pre-comment-fix-v1/xlsx_managed_compression_allocation.rs
- Previously reviewed SHA-256: 1072dadd121c7536cf0a2d8a98f795c408b61f261153066f241cfd4f3c2b0b36
- Final source SHA-256: adb0ddfb59691c393ace3b423b003b650526ee6a5ff4d2eed9b59e29bf85e907
- Final status: approved for the already-authorized isolated build; no workload or build was run by this review.

The final source preserves the reviewed drain and allocation evidence boundaries. Calibration requires the fixed probe count, Threads==1, measured counters, no allocator callbacks, unchanged live bytes, unchanged region peak, and unchanged process peak. The four native allocator phases and the fifth post-core worker-drain segment are checked for adjacent live-byte boundaries and counter conservation. The report retains both the four-phase core peak and the setup-relative peak including the drained core. Correctness gates remain outside the allocator regions and require changed-cell count, OutputBytes/CpuTasks accounting, bounded sink behavior, exact output and semantic hashes, exact inverse restoration, unchanged source hash, and clean managed resources before postdrop.

The only final revision is a comment correction in `ThreadObserver::drain_to_one`: it now states that pre-sample calls are outside operation regions and the post-core call is the measured worker-drain segment, without native timing attribution. An exact unified diff against the retained prior source was generated with:

    diff -u "/home/zhuhe/litchi-performance-0957/revisions/pre-comment-fix-v1/xlsx_managed_compression_allocation.rs" "/home/zhuhe/code/litchi/tools/perf-baseline/src/xlsx_managed_compression_allocation.rs"

The diff has 13 physical lines including headers and changes only the two comment lines at the method documentation; no executable Rust token changed. The prior source remains retained at the path above. No comparative capture, build, or test result is claimed by this review.
