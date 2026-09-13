# 0546: retain XLSX shared traversal with exact event admission

The retained OOXML/XLSX path shares worksheet parsing between independent raw
and validation consumers, preserves authoritative fallback and validated-EOF
handling, and counts the exact conservative event bound with sparse skipping
and bounded single-byte marker counts. Limits and public APIs are unchanged.

Across both repeats of medium and dense-sparse primary workflows, planning p50
improves 24.77–26.51%, total workflow p50 7.09–10.80%, and total mean 7.02–10.31%.
Owner-scoped planning instruction counts fall 21.81–22.36%. Ordinary planning
allocated bytes fall 0.078–0.104%; incremental peak rises only 0.0025–0.0062%.
All frozen native, allocation, refusal, cap, profile and eager gates pass.

The ten sparse/cap planning rows pass the 5% regression envelope, with the
largest p50/mean increase 2.66%. Synthetic sparse-comment controls improve
88–90%; this is a planning-only control result. The isolated clustered scanner
still regresses about 174% (about 28 microseconds), so the primitive is not
universally faster. Same-invalid refusal latency and allocations also regress;
the explicit candidate-invalid/baseline-valid envelopes pass at maxima 1.1021x
latency and 1.000034x incremental peak. These envelopes do not erase adverse
same-invalid rows. Whole-child RSS/counters remain diagnostics; page-fault drift
precludes a stable page-fault improvement claim.

All nine final quality commands pass, including 1,306 workspace tests,
all-features checking, warning-denied Clippy, rustdoc, formatting and crate
boundaries. Frozen source snapshots, raw receipts, ABBA samples, independent
source/results reviews, every adverse/drift row and deterministic analyzer
replays are retained. The analysis-path and formatting-only adaptations are
explicitly recorded. The temporary build tree is removed after executable
hash verification, and both evidence layers are sealed.

OLE2 and OOXML performance remain the priority. ODF is deferred until that
optimization goal completes; iWork remains excluded. Next measure fresh
post-change attribution, including the OLE2 CFB exact-chain loop opportunity;
its proposed terminal-proof fast path is design-only and has no speedup claim.

See [integrated results](integration/results-review.md), [decision](integration/decision.json), [OLE2 next opportunity](ole2-next-opportunity.md), and run `python3 -B docs/performance/results/change-0546/verify.py` for strict replay.

Whitespace checking covers production sources and edited reporting documents. Immutable raw logs, source snapshots and frozen review inputs preserve their original bytes; they are hash-verified rather than normalized. Outer diagnostic test/assembly stdout is losslessly gzip-compressed and verified against original receipt hashes. Precleanup attempts retain pending/missing-review-name findings; the verifier now reads the actual results-review.md artifact.
