# 0516 fourth-unit resource review

`performance_claim: none`

`claim_authorized: false`

This is a bounded, read-only resource review of the fourth-unit candidate
snapshot. The candidate source manifest SHA-256 is
`9127ae1530b3892f4d0f7b4cc7e608b0e678617827bf2b4754b5fca33e04fdc9`.
The review did not build, benchmark, or edit production source. It covers the
candidate's new output-fed worksheet parser and its final Store handoff. The
ordinary exact-output parser remains the fallback authority.

## Claim boundary

`SpeculativeBudget` in `raw/compact.rs:24-54` is a candidate-local admission
gate. It accounts for the compact output `Vec` capacity while the new
`EventParser` and its proof state coexist with that output. It does not claim a
128 MiB bound for the whole operation or process.

The existing compactor path remains a baseline owner. Its `NsReader`, namespace
resolver, `Writer`, `preserve: Vec<bool>`, normalized start-tag attribute
buffer, source events, and web probe are created by `changed_observed`
(`raw/compact.rs:275-381`) whether or not the provisional parser is retained.
Those allocations are deliberately outside the new threshold, as stated by
the candidate comment at `raw/compact.rs:24-28`. The output capacity is counted
because it is the exact byte result that coexists with the new parser; this is
not a whole-process or RSS claim.

The candidate now reads decoded parser fields after each event
(`raw/worksheet/codec.rs:1298-1316`). A shared-string cell or shared formula
sets `requires_exact_parser` and drops the feed before finalization. That is the
right boundary for the package shared-string callback and formula translation;
their allocations belong to the ordinary exact parser path and are not
silently claimed by this 128 MiB candidate proof.

## Findings

### Merge finalization remains undercounted

The current `MERGE_INDEX_BYTES` is four `Rect` slots plus sixteen `u32` slots,
128 bytes per merge (`raw/worksheet/codec.rs:1255-1279`). It covers a useful
part of the exact arrays, but it does not cover the actual peak in
`merge::Index::new`:

* `indices`, `nodes`, and `spans` are allocated at
  `merge.rs:107-134` (4, 20, and 4 bytes per entry on the current layouts);
* `build` allocates lower, spanning, and upper partition vectors while the
  tree arrays remain live (`merge.rs:176-273`); and
* validation keeps an active `BTreeMap` and an expiry `BinaryHeap` live
  (`merge.rs:294-334`). The heap can also retain its old allocation while a
  growth allocation is installed. The map's node allocation and allocator
  rounding dominate the explicit fixed-width arrays.

The existing parser charge for `parser.merges` (`codec.rs:1170-1202`) covers
the raw `Vec<Rect>` geometric capacity, but not those final-index and
validation owners. A practical admission allowance is a fixed 64 KiB base
for container startup and one **1 KiB per merge** for final index/validation
scratch. The 1 KiB term leaves room above the approximately 28 bytes per merge
for the three final arrays, roughly 32 bytes for transient partition/index
vectors, roughly 32–48 bytes for heap capacity and growth overlap, and a
conservative several-hundred-byte allowance for a `BTreeMap` entry/node and
allocator rounding. This is an admission bound for the current target and
should be replaced by exact capacity accounting if a cross-allocator proof is
required. The current 128-byte term is therefore an underbound; it should be
replaced by a fixed-base plus 1 KiB/merge term, accepting earlier fallback for
merge-heavy worksheets.

### Seen-row accounting is now explicit and practically sufficient

`Parser::seen_rows` is a `HashSet<u32>` (`raw/worksheet/model.rs:161-181`) and
is populated for every row (`raw/worksheet/codec.rs:589-602`). The candidate
now adds `SEEN_ROW_BYTES = 16 * size_of::<u32>()` to `ROW_BYTES`
(`raw/worksheet/codec.rs:1162-1169`). This is 64 bytes per retained row. A
`HashSet<u32>` stores a four-byte key plus control metadata; allowing a full
resize overlap and allocator/alignment overhead remains well below 64 bytes
per row on the current hash table implementation. The first small table is
also covered by the first row's term, while the `HashSet` wrapper itself is in
`size_of::<Parser>()`.

Thus the explicit 64-byte term is a generous practical amortization and is no
longer a missing owner. For a proof that must survive a changed standard
library layout, add a small fixed hash-table base (for example 1 KiB) or remove
the set using the parser's already enforced sorted-row invariant. The current
term is acceptable for this snapshot's bounded admission model.

### Plain `str` values still need a final payload allowance

The candidate uses a three-copy event scratch term, but retained payload is
still multiplied by only two (`raw/worksheet/codec.rs:1152-1158,
1197-1229`). Shared-string cells are excluded as described above. A normal
`t="str"` cell remains eligible, however. During `Parser::finish_store`, raw
cells stay live while `materialize` runs (`raw/worksheet/codec.rs:410-432`),
and `parse_value` decodes the raw value into a new `String` before converting
it into `Text(Arc<str>)` (`raw/worksheet/semantic.rs:124-132`). The raw encoded
string, decoded string, and Arc-owned string can coexist. The final estimate
currently adds only fixed `Stored` overhead and no dynamic text bytes
(`raw/worksheet/codec.rs:1248-1280`); event scratch has already been released.

This is a finalization undercount for eligible plain `str` values. The minimal
candidate-local correction is to raise the retained payload multiplier to
three, or add an equivalent dynamic allowance specifically for `str` material-
ization. If String geometric capacity is included in the proof rather than
treated as a separate allocator margin, charge the corresponding reserve
factor as well.

### Other terms are deliberately conservative

The per-cell `PendingCell` charge is applied once per retained cell even though
only one pending cell is live. The materialized-cell and row terms are doubled,
and `CELL_ROW_INDEX_BYTES` is charged per cell although `Store` creates one
cell-row entry per distinct row (`cell.rs:763-789`). `COLUMNS_BYTES = 4 MiB`
also covers the fixed column assignment tree and its range conversion with
substantial margin. These are overbounds that help the fallback decision and
do not repair the missing merge finalization or eligible `str` payload terms.

## Dense fixture

The candidate's 256-by-256 dense numeric fixture is MCE/x14ac-free and uses no
shared strings, shared formulas, or merges. Its retention assertion is in
`raw/compact_output_tests.rs:231-251`. It fits the default 128 MiB admission
ceiling with the current terms, so the normal dense numeric case qualifies.
That result does not extend the claim to merge-heavy, text-heavy, shared-string,
or shared-formula worksheets.

## Disposition

The decoded-field exclusion closes the external shared-string and
shared-formula ownership gap, and the explicit seen-row term is practically
adequate. Before treating the candidate budget as a sound new-state proof,
replace the 128-byte merge-finalization term with a fixed-base plus
1 KiB/merge allowance and account for the third live payload copy of eligible
plain `str` values. No performance or whole-process memory claim follows from
this review.
