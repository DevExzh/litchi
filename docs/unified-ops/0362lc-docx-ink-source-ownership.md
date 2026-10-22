# 0362lc — Retain the existing OPC owner in Ink patches

- Parent: `e08101d127`, the matched 0362lb allocation review.
- Status: correctness validated; matched allocation and native evidence pending.

An Ink edit previously cloned the complete OPC package into its private
source state even though the DOCX package already owns that graph through
immutable shared storage. The new no-op durable Commit extended the lifetime
of this duplicate through fingerprinting, raising peak memory. Ink source
states now retain the existing shared OPC owner. Changed candidates still
detach before mutation; validated target states own their separate immutable
graph. The sharing accessor is crate-private, and every changed internal
consumer migrates together.

The new integration regression drops the producing owner, retains sibling
no-op and changed patches, replays them onto an independent matching copy,
rejects the no-op against a changed target, and restores the exact source with
the inverse. Publishing one result preserves both sibling guards and the
independent input. Existing lexical, payload, signature, policy and malformed
durable proofs continue to pass.

Validation passes all 58 Ink integration tests, 1,489 library tests and 170
doctests (31 existing ignored examples), all-target/all-feature Clippy with
warnings denied, and formatting. The
[receipt](results/0362lc/validation.json) records the five changed source
files, exact commands/log hashes and dependency lock. This private repair
does not change the public surface validated by 0362lb's glossary profiles.
The unchanged 0362l probe will compare the committed repair with the original
`d0406be017` baseline; neither earlier peak regression is accepted.
