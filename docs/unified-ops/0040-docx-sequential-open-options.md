# 0040 — Managed DOCX sequential ingress

Lazy DOCX Package now exposes from_reader_with_options and canonical from_bytes /
from_bytes_with_options. The old source-backed from_reader_with_limits method
is deleted without an alias; active SVG callers and both tail-append fuzz targets
use OpenOptions. Both DOCX glossary profiles ban the retired method.

Source limits alone preserve the existing OPC move-owned input path and its
immutable-source memoization. Explicit cache, read-ahead or execution settings
use a bounded reader adapter and then the same positional options path as 0039.
Ordinary payloads remain cold until queried. Borrowed byte input is copied;
no zero-copy sharing claim is made.

## Accounting and lifetime

SourceLimits::InputBytes bounds the compressed archive length, including the
usual one-byte probe at the ceiling. It is distinct from the execution budget's
Resource::InputBytes: that cumulative resource charges accepted sequential input
and subsequent OPC positional reads independently. Adaptive read windows fit
remaining input capacity. Even EOF needs one byte of preflight input headroom,
but a zero-byte result commits no input. Exhausted input budget refuses before
calling the reader; it never grants an uncharged probe.

Each physical read attempt consumes one Work unit after input admission. Every
Interrupted retry checks cancellation and consumes another unit. Copy work is
also charged. Overreported read counts are rejected before committing input;
accepted counts commit before a post-read cancellation check. Partial I/O,
cancellation and limit errors preserve cumulative consumed input while releasing
all transient/retained memory from failed ingress.

Managed memory includes an 8 KiB scratch buffer, retained vector capacity and
source-owner metadata. Growth reserves the complete replacement while the old
buffer remains live, checks the reported allocation capacity, and frees the old
buffer before releasing its charge. The source ceiling also bounds capacity.
Scratch is released before OPC admission. The final ReadAt wrapper holds the
buffer and owner reservations beside OwnedSource, preserving the core's global
source identity. Package clones, source artifacts and retained OPC snapshots
therefore retain the charge; dropping the last source releases it. This is
logical resource accounting with ordinary RAII destruction, not an allocator-
total or process-RSS bound.

Cache/read-ahead/context options propagate through the new reader and byte
routes. Raw source export retains its existing behavior; package publication
switches read-ahead back to exact reads and releases the window, while the
sequential input owner remains charged until its final handle is dropped.

## Verification

Eight private ingress tests cover exact/one-under growth peaks, final-owner
lifetime, exact and one-over source limits, adaptive input admission and EOF
headroom, bounded interrupted retries, cancellation during a read, overreporting
(including the ceiling probe), refusal before reader I/O, and partial I/O failure.
Four integration tests cover managed source/artifact/clone lifetime, independent
archive-length versus execution-input accounting, read-ahead publication,
cleanup after failed OPC admission, and default/cache-only byte parity and limits.

The new sequential_ingress fuzz target varies short reads, interrupted retries,
overreported counts, cancellation and read-ahead under finite source/execution
limits. Its shared deterministic smoke driver crosses 256 modes with one valid
DOCX and three malformed inputs, asserting release of all memory charges.

- Full-feature DOCX/operations unit and integration suites: 2,426 passed,
  one existing ignored (2,037 DOCX and 389 operations).
- Minimal-feature ingress unit suite: eight passed.
- Warning-denied DOCX doctests: 85 passed, 31 existing ignored.
- Workspace all-feature library compilation, including facade/Python: passed.
- Warning-denied DOCX/operations library Clippy: passed.
- Glossary: 11 profiles, 140 banned-name rules, zero complete families; passed.
  All 21 checker tests pass.
- Crate boundaries: 66 packages, 267 internal dependencies and 11 existing
  explicit debt items; passed.
- All DOCX fuzz bins compile. Deterministic smoke passed: 64 valid modes
  opened, 192 valid-input modes refused overreporting/cancellation, and all
  768 malformed-input modes refused. No memory charge remained in any case.
  Reproduce with `CARGO_TARGET_DIR=target cargo run --manifest-path crates/litchi-docx/fuzz/Cargo.toml --bin smoke_sequential_ingress`.
- Independent read-only review checked ownership, admission and failure paths.
  Its Work-accounting finding was corrected: refused input admission is not a
  physical read and consumes no per-read Work unit. The final full suite and
  smoke run include that correction.
- Changed Rust formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

## Remaining work and cleanup

Owning DOCX still explicitly refuses lazy resource policies. Canonical owner
consolidation, semantic limits, validated crypto policies, specialized transaction
migration, other formats, Python/CLI, parity and full acceptance remain active.
No performance improvement or sustained fuzz campaign is claimed.

The six base-checkout fixture copies listed with hashes in 0035, build output,
generated root/DOCX-fuzz lockfiles and temporary logs were removed before commit.
No worktree was created; base, iWork and existing worktrees are preserved.
