# 0548 CFB checkpoint candidate source note

This directory contains an isolated one-file candidate for the private
`SectorChainScratch::collect_exact` path in
`crates/litchi-cfb/src/file.rs`. The candidate is prepared against repository
revision `dbe847aba2be674892673137450e360000360175`; production sources were
not edited, built, tested, or committed while this candidate was prepared.

| Item | SHA-256 |
| --- | --- |
| Current production `crates/litchi-cfb/src/file.rs` | `72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b` |
| Candidate `candidate-sources/file.rs` | `c773ab45a0d6b0fe0918406432933e52f7e7e2432d0897c3a4d3f877081ec427` |
| `candidate.patch` | `4a4168a097d6490ba399b71101ecb05d3ede7dec29273ef26c8bf1c4163d8d65` |

The patch applies with `git apply --check` to the current checkout. It keeps
the existing preflight checks, exact `sector-chain entries` reservation,
`sector-chain map` preparation, and their order unchanged. `validate_stream_allocations`
and all callers are unchanged. No public API, dependency, unsafe code, or
resource limit changes are included.

## Candidate mechanism

After the established reservations and `prepare_visited`, the successful
walk uses private scalar Brent-style state with checkpoints at sequence
positions `0, 1, 3, 7, ...`. It retains the existing index conversion and
table-bound checks, appends in the same order, loads the same immutable table,
and performs the same intermediate and final marker checks. A checkpoint hit
is never returned as a new diagnostic. The speculative prefix is cleared and
`walk_prepared_authority` replays the original visited-bit walk, preserving
the exact error text and precedence.

Every fast structural failure and every checked counter failure uses the same
prepared replay. The outer `collect_exact` result handler still resets both
logical scratch lengths on every error. A terminal fast success is justified
by the reviewed immutable-table proof: an exact walk ending in `ENDOFCHAIN`
cannot contain a repeated sector. Such a success leaves the retained visited
words zero and only keeps `visited.bit_len == allocation_table.len()`; the
next call's existing `prepare_visited(...).fill(0)` restores the normal map
state. Focused tests assert this private state explicitly.

`power` is initialized to one and is updated with `checked_mul(2)`, clamped to
`expected_count`; an overflow selects the declared-count clamp rather than
wrapping. `distance` uses `checked_add(1)` and rejects a value above
`expected_count`, which requests authoritative replay. Thus checkpoint state
cannot wrap or grow beyond the declared work bound. The boundary test covers
power multiplication overflow, distance addition overflow, and the declared
count guard.

The replay helper is entered only after the two existing fallible reservations
and map preparation have succeeded. It clears the speculative sector prefix,
then pushes at most `expected_count` entries into the exact reserved
`sectors` capacity and sets visited bits in the already prepared map. It does
not reserve or resize either scratch buffer, so replay cannot introduce a new
scratch allocation, move an allocation-failure boundary, or change a resource
label. Existing formatted corruption errors may still allocate their error
text, as they do in the current authoritative path.

## Candidate-only verification coverage

The snapshot adds a Rust differential test that enumerates every allocation
table of lengths zero through four over all table slots, one out-of-range
slot, `ENDOFCHAIN`, and low/middle/high reserved markers. It tries the same
representative start values and counts through one above the table length and
compares the candidate with `collect_sector_chain_exact` for exact success
order and diagnostic text. It also checks reusable scratch capacity and
pointer behavior through repeated calls, clean sector/map state after every
error, zero speculative visited bits after terminal success, and the checked
counter boundary cases.

These tests are included for the root agent's candidate build and review. No
build or test command was run in this isolated preparation step.
