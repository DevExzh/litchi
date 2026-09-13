# 0549 CFB checked test-and-mark candidate source note

This directory contains an isolated one-file candidate for the private
`SectorChainScratch::collect_exact` path in
`crates/litchi-cfb/src/file.rs`. The candidate is prepared against repository
revision `6d9fbb7401729aacf2afc5d6b6c681a9e7384056`; the live production file
was not edited, built, tested, captured, or committed while this candidate was
prepared.

| Item | SHA-256 |
| --- | --- |
| Current production `crates/litchi-cfb/src/file.rs` | `72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b` |
| Candidate `candidate-sources/file.rs` | `711572fd5779aff9efdfd3e452040daa378e23aad22bb6ecb83841dbbe903377` |
| `candidate.patch` | `0602a50efe65c800ea36163d22af19d4f44dac9a77f4d87861230956a23fb40d` |

## Candidate mechanism

`CheckedBitSet::test_and_mark` combines the existing read and write at the
exact scratch collector's visited-bit update. It preserves both existing
fallible bounds diagnostics and the backing-word diagnostic, returns whether
the bit was already present, and uses one checked mutable word lookup. The
implementation loads `old_word`, computes `new_word = old_word | mask`, stores
the result, and returns `old_word & mask != 0`. Rewriting an already-set word
does not change observable scratch state; retaining the combined load/store
shape leaves the optimizer free to select a bit-test-and-set instruction.

The existing `contains` and `insert` helpers remain unchanged for their other
callers. Only `SectorChainScratch::collect_exact` uses the combined helper.
The duplicate check remains before the sector append, so the first cycle
diagnostic and error text have the same precedence. Preflight validation,
sector-vector reservation, visited-map preparation and zero-fill, reset on
success/error, table access, index checks, marker checks, allocation labels and
all callers remain unchanged. The helper adds no public API, dependency,
unsafe code or resource-limit behavior.

## Candidate-only verification coverage

Two private Rust tests directly cover absent and present bits at the first,
last and next word positions, the exact `bit_len` boundary, a missing backing
word, and unchanged pointer/capacity state after successful and failing calls.
The existing scratch differential and reset/oracle tests remain in the
candidate source. An exhaustive helper enumerates 376,264 table/start/count
combinations for table lengths zero through four, comparing exact sector order
and diagnostic text with `collect_sector_chain_exact` and checking populated
visited membership on successful non-empty chains.

Static preparation checks passed:

* `rustfmt --check --edition 2024 docs/performance/results/change-0549/candidate-sources/file.rs`
* `git apply --check docs/performance/results/change-0549/candidate.patch`

No build, Rust test, benchmark, allocation capture, profile, hardware-counter
capture or guard run was performed in this isolated preparation step. The root
agent must measure the candidate's generated assembly and run the complete
matched correctness, allocation, profile, guard and native gates before any
production-adoption decision.
