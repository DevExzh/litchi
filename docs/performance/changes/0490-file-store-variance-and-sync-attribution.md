# 0490: resolve file-store variance with controlled synchronization evidence

0489 retained a small file-store tail regression in one repeat. This follow-up
alternates retained before/after binaries across six blocks and validates all
72 processes, 4,320 samples and exact output identities. Normal file-store
median latency improves 6.73% on average, but tail intervals span both
directions and individual regressions remain. No consistent tail gain or
production change is claimed.

[Separate diagnostics](../results/change-0490/results-review.md) attribute about
79–82% of traced median operation time to the unchanged data-sync call. The
selected syscall counts match between versions. Data synchronization remains
required; its latency is a measured reason to move attention to broader
source-provider/cold coverage rather than removing correctness or durability
work. Twenty helper tests and both retained-data verifiers pass. Temporary
replay directories are cleaned, and the complete evidence is sealed.

The [next implementation contract](../results/change-0490/next-implementation.md)
uses the existing end-to-end DOCX full-text route for verified-cold/provider
coverage, keeping prepared warm queries ineligible and genuine borrowed input
explicitly unimplemented. Bounded concurrency, native producer coverage and
the full non-iWork goal remain open.
