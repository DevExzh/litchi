# 0548 candidate source review

This is a read-only review of the isolated 0548 CFB candidate. The candidate
source snapshot is
[`candidate-sources/file.rs`](candidate-sources/file.rs), SHA-256
`421a1c4fd23a4bd7571e71e032caa297f842c50509d0fd3ef1d9e80af2886156`. It is
based on repository revision `dbe847aba2be674892673137450e360000360175` and
the live production `crates/litchi-cfb/src/file.rs` SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The
prepared patch is `candidate.patch`, SHA-256
`f4cf091a2168b18a7b5d929f25b779db1f925104307e277e556f25c7dbe91945`.

No production source was changed by this review. I ran no Rust build, test,
profile, benchmark, or capture. The candidate remains isolated until the
coordinator completes the frozen baseline phase and applies this exact source
snapshot.

## Static disposition

The candidate is admitted to the isolated baseline/candidate measurement
campaign, subject to the wording correction and the runtime gates in
`plan.json`. I found no source-level correctness blocker in the checkpoint
walk, its prepared replay, or its interaction with physical-sector claims.
This disposition does not retain the candidate in production: retention still
requires the exact candidate source and binary to pass the differential/error
contract, malformed-input guard, allocation and memory gates, OLE2/OOXML
workflow gates, profile attribution, and final quality checks.

The source comment on `walk_prepared_authority` says that replay performs “no
fallible allocation.” The implementation does avoid `try_reserve`, `resize`,
or any new scratch-buffer growth after the established reservations, so it
preserves the allocation boundary relevant to this experiment. Its formatted
`OleError::CorruptedFile` branches can still allocate the error `String`, as
the existing authoritative walk does. Reports should therefore say “no new
fallible scratch-buffer allocation or growth during replay,” and the source
comment should be narrowed to that meaning before final retention.

## Changed boundary and terminal proof

The only candidate production boundary is the private
`SectorChainScratch::collect_exact` at candidate lines 2935–3062. Its callers
and publication path are unchanged: `validate_stream_allocations` owns one
scratch value for MiniFAT and one for regular FAT, consumes only
`sectors()`, and then performs the existing mini-sector or physical-sector
claims. The candidate does not move claims, change role reconciliation, expose
sector IDs, or alter a public API.

For a nonempty call, the source retains the established order:

1. clear the prior logical sector and visited lengths;
2. reject an invalid empty declaration, start marker, or declared count;
3. reserve `sectors` with the existing `"sector-chain entries"` label;
4. prepare and zero the visited map with the existing `"sector-chain map"`
   label; and
5. walk exactly the declared number of positions.

The fast walk still converts and bounds-checks the current sector before
checkpoint comparison, appends it after that comparison, loads the immutable
allocation-table successor, and performs the existing intermediate and final
marker checks. It omits only the per-position visited load/mask/store on the
ordinary path. A terminal fast success requires the final successor to be
`ENDOFCHAIN`.

Let `N` be `expected_count`, `s_i` the sector at declared position `i`, and
`f(s)` the immutable allocation-table entry for a valid sector. Every
non-final fast position checks that its successor is neither `ENDOFCHAIN` nor
a reserved marker; every later position checks the current index before using
it. At the final position the successor must be exactly `ENDOFCHAIN`.

If a sector repeated, `s_i == s_j` for `i < j`, immutability would give
`f(s_i) == f(s_j)`. The suffix would consequently remain periodic and could
not terminate at `ENDOFCHAIN`; if the repeated position were final, its
successor would simultaneously have to be the earlier nonterminal successor
and `ENDOFCHAIN`. Thus an exact walk that reaches an accepted terminal
successor contains no repeated sector. This justifies leaving the visited
words zero on the fast success path. The proof relies on the borrowed table
remaining unchanged for the call, which is true for the private callers and
the `&[u32]` input.

The proof establishes only the exact terminal case. It does not make the
checkpoint state an authoritative cycle diagnostic and does not justify
skipping the visited map on malformed walks without replay.

## Checkpoint and bounded malformed work

`SectorChainCheckpoint` uses scalar state only: a `u32` checkpoint, a positive
power-of-two block length, and a distance. Its replacement positions are
`0, 1, 3, 7, ...`. The current sector is compared with the prior checkpoint
before append and table lookup; after the successor has passed the same
marker checks as before, `advance` updates the distance and replaces the
checkpoint at the end of the current block.

When a cycle has a noncyclic prefix of length `mu` and period `lambda`, the
first checkpoint at or after the prefix eventually lies in the cycle. Its
comparison window has the corresponding power-of-two length. Once that
window is at least `lambda`, a later position in the window equals the
checkpoint and requests replay. If the declaration ends first, a final
non-`ENDOFCHAIN` successor requests replay. Therefore a short cycle cannot
silently consume the whole declared count merely because it was not equal to
the first checkpoint. The exact constant factor is a measurement and proof
claim for the campaign, not an assumption made from the type names.

The arithmetic is checked. `distance.checked_add(1)` returns false rather
than wrapping, the distance is rejected if it exceeds the declared count,
and checkpoint power doubling uses `checked_mul(2)` and clamps to
`expected_count` on overflow. A false arithmetic result requests the same
prepared authoritative replay. In the production state initialized by
`Default`, the loop can perform at most `N - 1` advances and the distance and
power remain within the declared count. The focused boundary test also
exercises synthetic near-`usize::MAX` states; that test is useful evidence for
the guard branches but is not a proof that a host can allocate a table of that
size.

The checkpoint is deliberately conservative about diagnostics. A match does
not return a new error. It clears the speculative vector and invokes
`walk_prepared_authority` from the original start, so the earliest current
cycle or any earlier structural error remains authoritative. Every fast
structural failure follows the same replay path. This includes conversion or
table-bound failure, an impossible table lookup after the bound check, an
early terminal marker, a reserved marker, a late nonterminal successor, and a
counter guard failure.

## Error precedence

The replay preserves the original diagnostic sequence:

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

Preflight failures return before either reservation. The fast walk performs
the index conversion and table bound check before checkpoint comparison, just
as the old walk performed them before its visited test. A checkpoint hit is
replayed from the beginning, so a prior cycle wins over a later late-length
diagnostic. A cycle cannot create a genuinely new early terminal, reserved
marker, or out-of-range successor after its first occurrence because the
allocation table is immutable, but replaying those failures as well keeps the
implementation's precedence uniform and conservative.

`walk_prepared_authority` reproduces the old visited check, insert, append,
successor load, and marker branches. Its `map_err` and formatted messages use
the same text and table-name interpolation as the original method. The outer
result handler still calls `reset()` for every error, clearing both logical
lengths while retaining capacities. A replayed error can leave visited bits
set in retained words before that reset, matching the old error-path
possibility; the next call's `prepare_visited(...).fill(0)` clears all retained
words before any use.

## Allocation ordering and replay state

The candidate keeps the `sectors` reservation before `prepare_visited`, with
the existing resource labels and the existing “declared count exceeds table”
preflight. It does not replace either fallible reservation with a lazy or
per-step reservation. `walk_prepared_authority` is entered only after both
boundaries have succeeded; it clears the speculative prefix and pushes at
most `expected_count` entries into the already exact-reserved vector. The
visited map was already sized and zero-filled, and `CheckedBitSet::insert`
only checks and updates an existing word. Consequently replay introduces no
new scratch-buffer reserve, resize, or allocation-failure boundary and cannot
relabel a reservation failure. The formatted corruption error itself may
allocate its message, as described above; that is ordinary error construction,
not scratch growth.

On a fast terminal success, `sectors` contains exactly the old sector order,
`visited.bit_len` remains the table length, and all retained visited words are
zero. The only current consumers of `SectorChainScratch` read `sectors()`;
they do not inspect the visited words. Each subsequent collection starts with
`reset()` and, for a nonempty call, a full `prepare_visited(...).fill(0)`. The
two table namespaces use separate scratch instances. This makes the changed
private successful-map postcondition safe for current consumers, while the
candidate's explicit zero-bit assertion prevents future readers from
mistaking an uninitialized map for a populated ownership set.

Physical claims remain after collection and use the exact returned sector
slice. A failed collection resets its scratch logical state before the error
escapes; no new partial chain is passed to `claim_chain`. No source limit,
integer limit, dependency, unsafe block, lock, provider, public type, or
ownership publication rule changes in this patch.

## Rust test coverage and limits

The candidate adds meaningful private Rust coverage:

* a differential matrix compares scratch results with
  `collect_sector_chain_exact` for valid, empty, cycle, invalid-start,
  invalid-index, early-end, late-end, invalid-marker, short-table, and
  empty-nonterminal cases, including exact error strings and repeated scratch
  reuse;
* an exhaustive test enumerates allocation tables of lengths zero through
  four over each valid slot, an out-of-range slot, `ENDOFCHAIN`, and several
  reserved markers, then tries representative starts and counts through one
  above the table length; it checks sector order, diagnostics, reset state,
  and zero visited words after terminal success;
* buffer pointer/capacity reuse and growth-before-walk behavior remain
  asserted; and
* a focused counter test covers checked power multiplication, checked
  distance addition, and the declared-count guard near integer boundaries.

The exhaustive test is an actual Rust differential test of this source, not
the semantic model from the preceding design work. It still cannot cover all
table lengths, allocator failure injection, a 32-bit `usize` target, or the
physical-sector claim graph. Those boundaries are covered by source-level
invariants and must be exercised or bounded by the coordinator's candidate
quality, allocation, malformed-input, and full crate test gates. The exhaustive
model and any proof receipt must remain labeled as semantic support rather
than native allocation or latency proof.

The measurement tooling must attribute the exact `collect_exact` symbol. The
new cold `walk_prepared_authority` helper contains the visited operations that
are intentionally absent from the valid fast path; substring matching on
`collect_exact` or aggregation of all similarly named symbols could charge
that helper to the hot collector and invalidate the proposed instruction
comparison. Assembly and profile review should therefore match the exact
final symbol name and keep the helper's cold call edge explicit.

The candidate is consequently suitable for the frozen two-repeat campaign.
Runtime retention remains conditional on every gate in `plan.json`; OLE2 and
OOXML remain the active priority, with ODF deferred until that goal completes.
