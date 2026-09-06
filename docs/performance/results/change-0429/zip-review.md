# ZIP short-read correctness review

`ZipEntries::next_entry` previously refilled even when `buffered == 46` and
requested 46 bytes even when part of that header was buffered. Valid capped
reads could therefore return `BufferTooSmall`; a final exactly buffered fixed
header could be skipped. Exhaustion with a partial header could return `None`.

The bounded change refills only when fewer than 46 fixed bytes are buffered,
requests the missing count, and retains the existing larger prefetch capacity.
A declared directory tail too short to complete the header returns `Eof`.
An empty tail still returns `None`. Fixed signatures, variable metadata,
ZIP64 extra decoding and cumulative metadata admission remain in their existing
paths. ZIP32 and ZIP64 directory end positions still exclude their terminal
records. No entry-count shortcut is added.

Four generated integration regressions fail before and pass after the fix.
Expanded coverage has six tests: a 1..128-byte read-cap/scratch matrix, exactly
buffered fixed records with and without variable fields, truncated trailing
records, and retained ZIP32/ZIP64 fixtures with extra fields. Full ZIP tests
pass 439 tests with two ignored before the two extra test cases were added;
the final six focused tests pass. ZIP strict Clippy and warning-denied docs
pass. OPC/PPTX consumers pass 1,290 tests with three ignored. The original
compile and lint failures remain in the command receipts.

A source-only independent reviewer confirmed the missing-byte calculation,
exact-buffer boundary, typed truncation error, and ZIP64 end-position behavior.
This is a correctness enabler for caller short reads, with no performance claim.

The existing ZIP fuzz target also passes a bounded 1,000-run smoke with seed
429, a 1 MiB input cap, and AddressSanitizer under Rust 1.98.1. Thirteen retained
ZIP32/ZIP64 seeds and exact copied crate/fuzz sources are hash-bound in
`checks/fuzz-source-custody.json`. The isolated build and generated lock stay
outside the workspace; the lock bytes are retained as `.lock.txt` evidence.
Cleanup verifies copied inputs and removes only the exact task directory.
This bounded fuzz smoke is not an exhaustive adversarial-input proof.
