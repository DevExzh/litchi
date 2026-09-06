# 0450 validation and scope

Final release evidence: 455 ZIP tests, 436 OPC tests and 59 PPTX source-backed
cross-copy/adversarial tests pass. The ZIP suite has five new tests. Two ZIP and
one OPC ignored tests are inherited. Strict ZIP all-target/all-feature Clippy and
strict fuzz-target Clippy pass without diagnostics. ZIP all-target/all-feature
check, warning-denied rustdoc, no-iWork workspace/all-target check, boundaries and
explicit-file formatting pass. The workspace tests completed before the fuzz-only
source edit; the portable verifier proves every other source remained identical.

Ten deterministic I/O cases compare independently indexed archives, using Store
and Deflate at 0/1/65,536/262,181/1,048,613 decoded bytes. The generator is retained
in the exact tested Rust source, whose SHA-256 is bound by the test receipts and
source manifest; it uses xorshift32 seed 0x5a17932d and the locked ZIP writer. The
machine-readable lines in the immutable test log retain exact calls/returned bytes.
These are structural correctness/I/O cases, not benchmark latency samples, and
no p50, 5% latency-regression clearance or allocation-peak claim follows. Existing
expected-byte-path latency still needs matched measurements during OPC adoption.

Each case compares actual decoded bytes and the complete compressed token with
the prior cold-read/expected-byte-capture path. The token remains usable after
source drop, appends through PreservationPlan, leaves an existing member readable,
and records zero generated Deflate payload bytes. Other new tests cover 3-byte
short reads, immediate compressed/decode callback failure, wrong CRC/decoded size,
truncated/trailing Deflate, actual CRC under zero-CRC compatibility, ZIP64, invalid
entry ordinal, source transport failure and capacity-overflow refusal before any
callback. No partial token or decoded output is returned on error.

The existing bounded parse_zip fuzz target now calls the combined path before
requiring a successful ordinary decode. On success it compares decoded bytes
with the borrowed reader. Limits remain 1 MiB input/entry/aggregate decoded bytes,
256 entries, bounded metadata and eight progress events per capture attempt.
The isolated Rust 1.98.1 build uses RUSTC_BOOTSTRAP=1, AddressSanitizer and sancov
instrumentation. Thirteen retained ZIP32/ZIP64 descriptor seeds and 1,000 runs
with seed 450 pass. This is a smoke run, not exhaustive or native compatibility
coverage. Its generated lockfile and complete copied-source/seed custody are kept.

Failures and corrections are retained, not overwritten. The initial compile
used a writer method on the wrong wrapper; candidate-zip-tests records that
failure and compile-draft-office.rs.txt retains its source. The first successful
comparison reused a strict-layout cache; candidate-zip-tests-r2 and its draft stay
historical. R3 uses fresh indexes and adds transport/allocation tests. A scripted
writer-roundtrip edit failed to find its formatted marker without changing source;
then a wrapper-type mismatch failed candidate-zip-strict-r2. The final corrected
source, final-zip-tests and candidate-zip-strict-r3 are authoritative. No output or
threshold was rewritten to hide a failed check.

No production OPC/PPTX caller adopts the new primitive in 0450. The next integration
must retain OPC source authorization, cache/single-flight behavior, semantic checks,
combined reservations and publication/token lifetime. Native breadth, cold I/O,
bounded append, repackaging and scaling remain open; the full non-iWork goal stays
active. Existing sealed bundles and user GOAL.md are preserved.

Portable replay passes from exported copies before and after cleanup. Seven
precleanup and eight postcleanup corruption probes reject, including derived I/O,
source custody, fuzz run configuration, false adoption, copied-source identity,
cleanup failure and unsealed members. Probe copies remove themselves before the
receipts are published. The isolated fuzz workspace was removed after sealed
precleanup replay, with 673 regular files totaling 411,782,557 bytes inventoried.
Both Cargo targets retain their original directory identities; the user GOAL.md
retains its pinned SHA-256 and remains untracked/unstaged. No other temporary
workspace or binary is owned by this batch. All CPU jobs are terminal.
Final seal covers 92 members plus SHA256SUMS, for 93 bundle files.
