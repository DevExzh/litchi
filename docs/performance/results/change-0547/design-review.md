# 0547 CFB terminal-proof chain-walk design review

This is a design-only review for the next OLE2/CFB opportunity. The current
0547 plan freezes a fresh baseline attribution of
`SectorChainScratch::collect_exact`; it does not admit a production candidate.
ODF remains deferred until the OLE2/OOXML optimization goal is complete, and
iWork is outside this review.

The reviewed production source is `crates/litchi-cfb/src/file.rs` at revision
`2118c6fb1d023005e224aaa938e84d6ed0588b70`, with SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The
relevant source boundary is the private `SectorChainScratch` at
`file.rs:2764` and its `collect_exact` method at `file.rs:2799`. This review
made no Rust, test, build, profile, capture, or commit change. The companion
[`proof_check.py`](proof_check.py) is a small semantic model and must not be
read as Rust allocation or latency evidence.

## Current contract

`validate_stream_allocations` owns two independent scratch objects. MiniFAT
streams call `mini_scratch.collect_exact`, regular streams call
`regular_scratch.collect_exact`, and only a completed sector vector is passed
to mini-sector ownership or `claim_chain`. The root mini-stream and the
general owned-result helper use separate collectors. A candidate cannot move
ownership publication, physical-role reconciliation, or the two table
namespaces into this experiment.

For a nonempty call, the current method does the following in this order:

1. Reset the prior sector length and visited logical length.
2. Reject an invalid empty declaration, invalid start marker, or a declared
   count larger than the allocation table.
3. Reserve `sectors` under `"sector-chain entries"`.
4. Prepare the visited map under `"sector-chain map"`: retain or grow its
   words, set `bit_len` to the table length, and fill every retained word with
   zero.
5. For each declared position, check `u32` to `usize`, check the table bound,
   test and set the visited bit, append the sector, load its successor, and
   check the final or intermediate marker.
6. Reset the output and visited logical length after every error. On success,
   leave the exact sector order and `visited.bit_len == table.len()`.

The exact diagnostics are part of the existing differential contract:

```text
Empty {table_name} chain must start with ENDOFCHAIN
Invalid start marker for {table_name} chain
{table_name} chain length exceeds its allocation table
Invalid sector index {sector} in {table_name}
Cycle detected in {table_name} chain at sector {sector}
{table_name} chain exceeds its declared length
{table_name} chain ends before its declared length
Invalid sector marker 0x{next:08X} in {table_name} chain
```

The checked sector/index error precedes cycle detection. Cycle detection
precedes appending, table lookup, and final/intermediate marker errors. The
fallible reservations precede all loop work. These boundaries must stay
unchanged even if the successful loop stops touching the visited words.

## Exact terminal proof

Let `N` be `expected_count`, let `s_0` be the validated start sector, and let
`f(s)` be the immutable allocation-table entry at sector `s`. The loop only
defines `f(s_i)` after checking that `s_i` is an in-range table index. For each
`i < N - 1`, an `ENDOFCHAIN` or reserved marker is rejected; a value below
`MAXREGSECT` becomes `s_(i+1)`. At the final position the only accepted
successor is `ENDOFCHAIN`.

Suppose two visited positions repeat, `s_i == s_j` for `i < j`. Both entries
are valid table slots, so immutability gives the same successor at both
positions: `f(s_i) == f(s_j)`. Repeating that equality while the walk remains
in the declared range makes the suffix periodic with period `j - i`. It
cannot reach `ENDOFCHAIN` at a later final position. If `j` itself is the
final position, its successor would be `ENDOFCHAIN`, while the earlier
occurrence's successor was already accepted as a nonterminal sector; that is
also a contradiction. Therefore an exact-`N` walk whose final successor is
`ENDOFCHAIN` contains no repeated sector.

The proof depends on the table being unchanged for the duration of the call,
which the `&[u32]` input and this private method's ownership provide. It does
not remove any index conversion, table bound, marker, terminal, reservation,
or output-order check. It proves only the successful terminal case; it is not
a cycle detector for a malformed prefix.

## Smallest terminal-only replay experiment

A minimal work-elimination experiment can retain both existing reservations
and `prepare_visited` exactly as they are, then omit `visited.contains` and
`visited.insert` from the ordinary walk. The fast walk still checks the
current index, appends each sector, loads the successor, and applies the same
early, late, and invalid-marker branches. On an exact final `ENDOFCHAIN` it
returns the collected vector under the terminal proof.

Every other fast-walk failure must discard the prefix and invoke the existing
ordered bitset walk from the original `start_sector`, with the same table,
count, and table name. The replay result, including its error text, is the
only result exposed to the caller. This is necessary because an omitted cycle
check can let an earlier cycle coexist with a later structural failure:

```text
table = [1, 0, ...]
start = 0
N = 3
```

The fast walk reaches sectors `0, 1, 0` and can report a late excess error at
its final step. The current collector reports `Cycle detected ... at sector
0` at the third loop iteration. Replaying from the beginning restores that
precedence. For an immutable table, a duplicate cannot be followed by a new
early terminal, reserved marker, or out-of-range state: the first occurrence
has the same successor, and reaching the duplicate already proves that the
successor was a valid in-range sector. Those later structural errors are
therefore impossible after a true cycle; replaying every fast failure remains
the simplest conservative implementation, while the late nonterminal case is
the concrete precedence conflict.

Before replay, `sectors.clear()` is required; otherwise the authoritative
pushes would follow a prefix already appended by the fast walk. The current
`prepare_visited` has already zero-filled the retained words, and the fast
experiment writes no visited bits, so the authoritative replay starts with a
fresh logical map. A defensive second `fill(0)` is harmless but adds failure
path work; omitting it requires the source invariant that no fast-path code
touches those words. The outer error reset must still clear the sector length
and set `visited.bit_len` to zero.

On a terminal fast success, the sector vector is exact but the retained map's
words are all zero while `visited.bit_len` remains the table length. Production
code never reads that map after `collect_exact`; the caller reads only
`sectors()`, and the next call begins with `reset` followed by
`prepare_visited(...).fill(0)`. Thus no stale bit can cross a call or cross the
separate MiniFAT/regular-FAT scratch objects. This is nevertheless a changed
private postcondition. The existing tests inspect logical length, capacity,
and pointer reuse but not successful bits; a candidate must add an explicit
success-then-reuse guard and document this state, or preserve the old bits at
the cost of replaying the success path and losing the proposed saving.

The existing exact reservations remain important even though the map is not
used on success. Removing or deferring either one would change resource
labels, allocation failure order, and the established allocation envelope.
With the reservations retained, authoritative replay can push into existing
capacity without a new fallible boundary. This experiment removes per-step
visited work only; it does not remove map zeroing, map capacity, or any
ownership check.

## Terminal-only malformed-work counterexample

Let the table have length `L`, with `table[0] = 0` and arbitrary entries after
it, and call the collector with `start = 0` and `N = L`. The current collector
returns the self-cycle error after two loop iterations. The terminal-only fast
walk has no reason to stop at the second sector: it can walk all `L` declared
positions and then report a late excess error, followed by an authoritative
replay of two iterations. Its modeled work is therefore `L + 2` loop
iterations versus `2` for the current collector, a ratio of `(L + 2) / 2`.
The ratio grows with the declared/table length. The finite CFB ingress limit
prevents an infinite loop, but the allocation table can still be very large;
the limit is not a constant-factor malformed-latency guarantee.

This is a real hostile-input concern even though all work remains finite. A
terminal-only candidate must either accept and explicitly gate this regression
or add a cycle checkpoint. A broad claim that “replay makes malformed input
safe” is insufficient: replay restores the error, but it does not restore the
current early-stop work bound.

## Brent-style checkpoint alternative

A more promising bounded-work variant keeps the same reservations and zero
fill, but tracks only scalar cycle state in the fast walk. Use a validated
`Option<u32>` checkpoint, a power-of-two block length, and a distance within
the block. The checkpoint positions in the model are `0, 1, 3, 7, ...`
(`2^k - 1`); at each
validated current sector, compare it with the previous checkpoint before the
table lookup. After the current successor has passed its existing
intermediate marker checks, advance the distance and replace the checkpoint
at the end of its block. The current source order remains:

```text
index conversion and table bound
checkpoint equality
append and table lookup
final/intermediate marker checks
checkpoint counter update for the next valid state
```

A checkpoint equality is a proven repeated valid sector. The candidate still
must not return a new checkpoint error: it clears the prefix and replays the
authoritative collector so that an earlier cycle or any earlier structural
error wins with the existing text. Any invalid index, early `ENDOFCHAIN`,
invalid marker, or late nonterminal successor also invokes the same replay.
An exact terminal success remains justified by the deterministic terminal
proof, even if no checkpoint equality was observed.

The checkpoint schedule gives a simple conservative malformed bound. Let
`mu` be the noncyclic prefix length and `lambda` the cycle period, so the
first repeated sequence position is `mu + lambda`. Choose the first
checkpoint position `b = 2^k - 1` at least `max(mu, lambda)`. The schedule
has `b < 2 * max(mu, lambda) + 1`; its checkpoint is in the cycle, and at most
`lambda` subsequent comparisons find the same state before the block ends,
since that block has length `b + 1 > lambda`. If the declared count ends
earlier, the late nonterminal check requests replay sooner and obeys the same
bound. Thus the fast walk
detects the cycle in fewer than
`3 * (mu + lambda) + 2` loop positions. The authoritative replay takes at
most `mu + lambda + 1` positions, giving a conservative total below
`4 * (mu + lambda) + 3`. The exact constants are less important than the
change from declared-`N` work to a constant multiple of the earliest exposed
cycle. If a structural error occurs first, fast work plus replay is at most
twice the current error position.

This bound assumes checked counter arithmetic. `distance` must never be
incremented past its current block length, and doubling `power` must not use
unchecked `usize` multiplication. A safe implementation can clamp the block
length to the declared count after a checked/saturating update; a block that
cannot finish inside the current declaration needs no further boundary. The
distance update still needs a checked invariant, and an arithmetic guard must
fall back to the authoritative walk rather than manufacture a new
`OleError` or panic. If an overflow guard falls back to authority, its
worst-case path should be recorded as a separate extreme-size limitation; a
strict constant-factor claim must prove the clamped-counter variant itself.
The fast loop's existing `for index in 0..expected_count` remains the
authoritative declared-work bound.

Brent does add an equality, counter increment, and occasional branch to every
successful sector. It may therefore give up some of the gain from removing
the bitset word load/mask/set. The fresh 0547 attribution must decide whether
that tradeoff is worth measuring. The checkpoint strategy is a bounded-work
hypothesis, not a speedup result. The exhaustive model includes it to prove
semantic parity over a small domain, not to predict native code generation.

## Error and state matrix

| Input/result boundary | Fast walk action | Authoritative result required |
| --- | --- | --- |
| `N == 0`, start not `ENDOFCHAIN` | Return unchanged preflight error | Empty-chain diagnostic; no replay |
| Nonempty invalid start | Return unchanged preflight error | Invalid-start diagnostic; no replay |
| `N > table.len()` | Return unchanged preflight error | Allocation-table-length diagnostic; no replay |
| Current index conversion/bounds failure | Stop before checkpoint equality | Exact invalid-index diagnostic from replay; a true earlier duplicate cannot lead to an out-of-range successor |
| Checkpoint/duplicate sector | Stop before append/table lookup | Exact earliest cycle diagnostic from replay |
| Intermediate `ENDOFCHAIN` | Stop at existing marker branch | Exact early-end diagnostic from replay; a true earlier duplicate cannot reach this marker |
| Intermediate reserved marker | Stop at existing marker branch | Exact invalid-marker diagnostic from replay; a true earlier duplicate cannot reach this marker |
| Final non-`ENDOFCHAIN` | Stop at existing late branch | Exact late-excess diagnostic, or an earlier cycle diagnostic from replay |
| Final `ENDOFCHAIN` | Return exact sector vector | Terminal proof establishes no duplicate; no replay |

For all replayed failures, the prefix is cleared before replay and the outer
error reset leaves an empty sector vector and `visited.bit_len == 0` while
retaining capacities. For terminal success, the proposed fast state leaves
the exact sector vector, `visited.bit_len == table.len()`, and zero visited
bits. The caller and next-call reset behavior justify that private state only
after a focused candidate test and source review; the model records it
explicitly so this change cannot be mistaken for an unchanged bitset
postcondition.

## Exhaustive model and required gates

[`proof_check.py`](proof_check.py) compares the current ordered collector model
with terminal-only replay and the checkpointed walk. It enumerates every table
of lengths 0 through 4 over all valid slot values, one out-of-range value,
`ENDOFCHAIN`, and representative low/middle/high reserved markers. It tries
the same start-value classes and counts `0..table.len()+1`, compares exact
diagnostic strings and successful sector order, and checks clean failure state.
It also prints a self-cycle scaling example showing terminal-only declared-`N`
amplification and checkpointed bounded detection. Its scope is semantic: it
does not model Rust allocation failure, retained capacities, alias mutation,
or a hypothetical target where `usize` cannot represent a `u32`.

Before a production candidate, retain the current fresh baseline attribution
and map the actual `collect_exact` and visited operations. If the target is
credible, a candidate-only differential must cover exact terminal success,
empty chains, short chains, late excess, invalid markers, invalid indexes,
self-cycles, multi-node cycles, final-position duplicates, repeated scratch
reuse, and every exact error string. It must exercise the existing fallible
reservation labels and prove that replay cannot allocate after the established
boundary. Run the full native two-repeat OLE2/OOXML gate, malformed/refusal
latency controls, allocator and RSS gates, source/binary/assembly review,
quality checks, and workspace boundary checks. Ir reduction or a successful
model run alone cannot justify retention.

The current disposition is therefore: terminal proof is mathematically valid
for exact terminal success; terminal-only replay is semantically recoverable
but has an unacceptable declared-`N` short-cycle work amplification unless
explicitly bounded; the Brent checkpoint/replay variant is the smallest
additional state that plausibly restores an early-cycle work bound and is the
next candidate hypothesis if the fresh attribution supports its overhead.
