# 0362lc — Retain the existing OPC owner in Ink patches

- Parent: `e08101d127`, the matched 0362lb allocation review.
- Status: correctness and matched allocation checks pass; native timing pending.

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

The [matched receipt](results/0362lc/evidence.json) now verifies source
`76148455d5` against `d0406be017`. The unchanged probe, dependencies, lockfile
and release settings produced four protected binaries. Their six preflight
outputs match exactly. Four allocation legs agree for every route; each leg
uses 100 operations and retains zero bytes after the measured operations.

| Route | Allocation calls, base → repair | Allocated bytes, base → repair | Peak bytes, base → repair |
| --- | ---: | ---: | ---: |
| Forward | 307,900 → 300,900 | 36,779,900 → 36,025,800 | 82,323 → 79,232 |
| Inverse | 436,300 → 427,300 | 73,803,400 → 72,910,300 | 252,189 → 251,917 |
| No-op | 30,100 → 28,000 | 2,301,600 → 2,004,600 | 3,309 → 1,848 |
| Forward and write | 320,600 → 313,600 | 217,352,400 → 216,598,300 | 605,584 → 605,584 |
| Inverse and write | 443,100 → 434,100 | 170,786,600 → 169,893,500 | 592,774 → 592,774 |
| No-op and write | 30,200 → 28,100 | 2,660,800 → 2,363,800 | 3,592 → 3,592 |

This closes the allocation regressions returned in 0362la and 0362lb. No
native timing attempt has run for this source. Timing qualification and
integration remain required before acceptance.
